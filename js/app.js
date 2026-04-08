/* ============================================================
   DocuFill AI — Main Application Logic
   100% client-side: template parsing, AI inference, PDF export
   ============================================================ */

(function () {
    "use strict";

    // ── DOM refs ──────────────────────────────────────────────
    const $  = (s) => document.querySelector(s);
    const $$ = (s) => document.querySelectorAll(s);

    const keyStatus        = $("#key-status");
    const modelSelect      = $("#model-select");
    const uploadArea       = $("#upload-area");
    const fileInput        = $("#template-file");
    const demoTemplateBtn  = $("#demo-template-btn");
    const downloadDemoBtn  = $("#download-demo-btn");
    const templateInfo     = $("#template-info");
    const fileNameEl       = $("#file-name");
    const removeFileBtn    = $("#remove-file");
    const placeholdersEl   = $("#placeholders-found");
    const userContentArea  = $("#user-content");
    const processBtn       = $("#process-btn");
    const manualBtn        = $("#manual-btn");
    const progressBar      = $("#progress-bar");
    const outputSection    = $("#output-section");
    const downloadDocxBtn  = $("#download-docx");
    const editToggleBtn    = $("#edit-toggle");
    const editPanel        = $("#edit-panel");
    const mappingsEditor   = $("#mappings-editor");
    const applyEditsBtn    = $("#apply-edits");
    const documentPreview  = $("#document-preview");
    const manualModal      = $("#manual-modal");
    const manualFields     = $("#manual-fields");
    const manualCancelBtn  = $("#manual-cancel");
    const manualApplyBtn   = $("#manual-apply");

    // ── State ─────────────────────────────────────────────────
    let templateHtml       = "";
    let placeholders       = [];
    let currentMappings    = {};
    let hfApiKey           = "";
    let originalDocxBuffer = null;  // Store original .docx for faithful modification

    const apiKeyInput      = $("#api-key");
    const saveKeyBtn       = $("#save-key");
    const keyInputArea     = $("#key-input-area");

    // ── Load API key: .env → localStorage → show input ───────
    async function loadApiKey() {
        // 1. Try .env (works locally)
        try {
            const res = await fetch(".env");
            if (res.ok) {
                const text = await res.text();
                // Make sure we got an actual .env and not an HTML 404 page
                if (!text.trim().startsWith("<") && !text.includes("<!DOCTYPE")) {
                    const match = text.match(/^HF_API_KEY\s*=\s*(.+)$/m);
                    if (match && match[1].trim() && match[1].trim() !== "paste_your_huggingface_token_here") {
                        hfApiKey = match[1].trim();
                        setStatus("\u2705 API key loaded from .env", true);
                        updateButtons();
                        return;
                    }
                }
            }
        } catch { /* ignore */ }

        // 2. Try localStorage
        const stored = localStorage.getItem("docufill_hf_key");
        if (stored) {
            hfApiKey = stored;
            setStatus("\u2705 API key loaded from browser storage", true);
            updateButtons();
            return;
        }

        // 3. Show manual input
        setStatus("Enter your Hugging Face API key below", false);
        show(keyInputArea);
        updateButtons();
    }
    loadApiKey();

    saveKeyBtn.addEventListener("click", () => {
        const val = apiKeyInput.value.trim();
        if (!val || !val.startsWith("hf_")) {
            setStatus("\u26A0 Key should start with hf_", false);
            return;
        }
        hfApiKey = val;
        localStorage.setItem("docufill_hf_key", val);
        setStatus("\u2705 API key saved to browser storage", true);
        hide(keyInputArea);
        updateButtons();
    });

    apiKeyInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") saveKeyBtn.click();
    });

    // ── Helpers ───────────────────────────────────────────────
    function show(el)  { el.classList.remove("hidden"); }
    function hide(el)  { el.classList.add("hidden"); }
    function setStatus(msg, ok) {
        keyStatus.textContent = msg;
        keyStatus.className = "status-msg " + (ok ? "ok" : "err");
    }

    function extractPlaceholders(html) {
        const re = /\{\{([A-Z0-9_]+)\}\}/g;
        const set = new Set();
        let m;
        while ((m = re.exec(html)) !== null) set.add(m[1]);
        return [...set];
    }

    // Extract placeholders directly from .docx XML (handles headers, footers,
    // and Word's habit of splitting {{PLACEHOLDER}} across multiple <w:t> runs).
    async function extractPlaceholdersFromDocx(arrayBuffer) {
        const zip = await JSZip.loadAsync(arrayBuffer);
        const set = new Set();
        // Log ALL files in the zip for debugging
        const allFiles = Object.keys(zip.files);
        console.log("[DocuFill] All files in .docx:", allFiles);
        // Scan ALL xml files in the archive
        const xmlFiles = allFiles.filter(f => f.endsWith(".xml") && !f.includes("_rels"));
        console.log("[DocuFill] Scanning XML files:", xmlFiles);
        for (const fname of xmlFiles) {
            const xml = await zip.file(fname).async("string");
            // Extract ALL text from the file by concatenating <w:t> across <w:p> and also raw text
            let fileText = "";
            const paragraphs = xml.split(/<w:p[\s>]/);
            for (const para of paragraphs) {
                let paraText = "";
                const tRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
                let tm;
                while ((tm = tRe.exec(para)) !== null) paraText += tm[1];
                if (paraText) fileText += paraText + "\n";
            }
            if (fileText.trim()) {
                console.log(`[DocuFill] Text in ${fname}:\n`, fileText);
            }
            // Normalize smart/curly braces to standard braces
            fileText = fileText
                .replace(/[\uFF5B\u007B\u2774\uFE5B]/g, "{")   // various left braces
                .replace(/[\uFF5D\u007D\u2775\uFE5C]/g, "}")   // various right braces
                .replace(/\u00AB/g, "{{").replace(/\u00BB/g, "}}") // « » guillemet style
                .replace(/\u201C/g, "\"").replace(/\u201D/g, "\""); // smart quotes
            // Scan for placeholders (flexible: letters, digits, underscore, space, hyphen)
            const re = /\{\{([A-Za-z0-9_ -]+)\}\}/g;
            let pm;
            while ((pm = re.exec(fileText)) !== null) set.add(pm[1].trim());
        }
        console.log("[DocuFill] Placeholders found:", [...set]);
        return [...set];
    }

    // Convert a plain-text value to HTML, detecting numbered/bulleted lists
    function formatValueAsHtml(text) {
        const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        const lines = text.split("\n");
        const result = [];
        let inOl = false;
        let inUl = false;
        let paraLines = []; // accumulate non-list lines into a paragraph

        function flushPara() {
            if (paraLines.length) {
                result.push("<p>" + paraLines.join("<br>") + "</p>");
                paraLines = [];
            }
        }

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const olMatch = line.match(/^\s*(\d+)[.)]\s+(.*)/);
            const ulMatch = line.match(/^\s*[-•*]\s+(.*)/);

            if (olMatch) {
                flushPara();
                if (!inOl) { if (inUl) { result.push("</ul>"); inUl = false; } result.push("<ol>"); inOl = true; }
                result.push(`<li>${esc(olMatch[2])}</li>`);
            } else if (ulMatch) {
                flushPara();
                if (!inUl) { if (inOl) { result.push("</ol>"); inOl = false; } result.push("<ul>"); inUl = true; }
                result.push(`<li>${esc(ulMatch[1])}</li>`);
            } else {
                if (inOl) { result.push("</ol>"); inOl = false; }
                if (inUl) { result.push("</ul>"); inUl = false; }
                if (line.trim() === "") {
                    flushPara();
                } else {
                    paraLines.push(esc(line));
                }
            }
        }
        if (inOl) result.push("</ol>");
        if (inUl) result.push("</ul>");
        flushPara();
        return result.join("\n");
    }

    function fillTemplate(html, map) {
        let out = html;
        for (const [key, value] of Object.entries(map)) {
            const formatted = formatValueAsHtml(value);
            const hasBlockElements = /<(ol|ul|div|table|h[1-6])\b/i.test(formatted);

            if (hasBlockElements) {
                // If the placeholder is the main content of a <p> tag, replace the whole <p>
                // to avoid invalid nesting of block elements inside <p>
                const pRegex = new RegExp(
                    "<p[^>]*>\\s*(?:<[^>]+>)*\\s*\\{\\{" + key + "\\}\\}\\s*(?:<\\/[^>]+>)*\\s*<\\/p>",
                    "g"
                );
                if (pRegex.test(out)) {
                    out = out.replace(pRegex, formatted);
                } else {
                    out = out.replace(new RegExp("\\{\\{" + key + "\\}\\}", "g"), formatted);
                }
            } else {
                out = out.replace(new RegExp("\\{\\{" + key + "\\}\\}", "g"), formatted);
            }
        }
        // Replace any unfilled placeholders with visible markers
        out = out.replace(/\{\{([A-Za-z0-9_ -]+)\}\}/g,
            '<span style="background:#fecaca;padding:1px 4px;border-radius:3px;font-family:monospace;font-size:.85em">[MISSING: $1]</span>');
        return out;
    }

    // ── Enable buttons when ready ─────────────────────────────
    function updateButtons() {
        const hasTemplate = templateHtml.length > 0;
        const hasContent  = userContentArea.value.trim().length > 0;
        processBtn.disabled = !(hasTemplate && hasContent && hfApiKey);
        manualBtn.disabled  = !(hasTemplate);
    }

    userContentArea.addEventListener("input", updateButtons);

    // ── File Upload ───────────────────────────────────────────
    uploadArea.addEventListener("click", () => fileInput.click());
    uploadArea.addEventListener("dragover", (e) => {
        e.preventDefault();
        uploadArea.classList.add("dragover");
    });
    uploadArea.addEventListener("dragleave", () => uploadArea.classList.remove("dragover"));
    uploadArea.addEventListener("drop", (e) => {
        e.preventDefault();
        uploadArea.classList.remove("dragover");
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener("change", () => {
        if (fileInput.files.length) handleFile(fileInput.files[0]);
    });

    async function handleFile(file) {
        if (!file.name.toLowerCase().endsWith(".docx")) {
            alert("Please upload a .docx file.");
            return;
        }
        try {
            const arrayBuffer = await file.arrayBuffer();
            originalDocxBuffer = arrayBuffer.slice(0); // keep a copy
            const result = await mammoth.convertToHtml({ arrayBuffer });
            templateHtml = result.value;
            // Extract from raw XML to catch headers/footers + fragmented runs
            placeholders = await extractPlaceholdersFromDocx(arrayBuffer);
            // Fallback to HTML-based extraction if XML scan found nothing
            if (!placeholders.length) placeholders = extractPlaceholders(templateHtml);
            showTemplateInfo(file.name);
        } catch (err) {
            alert("Failed to parse the document: " + err.message);
        }
    }

    function showTemplateInfo(name) {
        fileNameEl.textContent = name;
        show(templateInfo);
        if (placeholders.length) {
            placeholdersEl.innerHTML = "<strong>Placeholders found:</strong> " +
                placeholders.map(p => `<span>{{${p}}}</span>`).join(" ");
        } else {
            placeholdersEl.innerHTML = '<em>No {{PLACEHOLDER}} tokens found. The AI will try to fill based on document structure.</em>';
        }
        updateButtons();
    }

    removeFileBtn.addEventListener("click", () => {
        templateHtml = "";
        placeholders = [];
        hide(templateInfo);
        fileInput.value = "";
        updateButtons();
    });

    // ── Demo Template ─────────────────────────────────────────
    demoTemplateBtn.addEventListener("click", async () => {
        const blob = await generateDemoDocx();
        const arrayBuffer = await blob.arrayBuffer();
        originalDocxBuffer = arrayBuffer.slice(0); // store for download
        const result = await mammoth.convertToHtml({ arrayBuffer });
        templateHtml = result.value;
        placeholders = extractPlaceholders(templateHtml);
        showTemplateInfo("demo-project-report.docx");

        // Pre-fill sample content for easy testing
        userContentArea.value = DEMO_CONTENT;
        updateButtons();
    });

    // ── Download Sample Template ─────────────────────────────
    downloadDemoBtn.addEventListener("click", async () => {
        const blob = await generateDemoDocx();
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "sample-template.docx";
        a.click();
        URL.revokeObjectURL(a.href);
    });

    // ── Progress ──────────────────────────────────────────────
    function setProgress(stepName) {
        show(progressBar);
        $$(".progress-step").forEach(el => {
            const s = el.dataset.step;
            if (s === stepName) {
                el.classList.add("active");
                el.classList.remove("done");
            } else if (shouldBeAfter(s, stepName)) {
                el.classList.remove("active", "done");
            } else {
                el.classList.remove("active");
                el.classList.add("done");
            }
        });
    }
    const STEP_ORDER = ["parse", "ai", "fill", "render"];
    function shouldBeAfter(step, current) {
        return STEP_ORDER.indexOf(step) > STEP_ORDER.indexOf(current);
    }
    function progressDone() {
        $$(".progress-step").forEach(el => {
            el.classList.remove("active");
            el.classList.add("done");
        });
    }

    // ── AI: Hugging Face Inference Providers (OpenAI-compatible) ──
    async function callAI(systemPrompt, userMessage) {
        const apiKey = hfApiKey;
        const model  = modelSelect.value;

        const body = {
            model: model,
            messages: [
                { role: "system", content: systemPrompt },
                { role: "user",   content: userMessage }
            ],
            max_tokens: 1500,
            temperature: 0.2
        };

        const res = await fetch(
            "https://router.huggingface.co/v1/chat/completions",
            {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${apiKey}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(body)
            }
        );

        if (!res.ok) {
            const errBody = await res.text();
            if (res.status === 401 || res.status === 403) throw new Error("Invalid API key. Ensure your Hugging Face token has 'Inference Providers' permission.");
            if (res.status === 503) throw new Error("Model is loading. Please wait ~30s and try again.");
            throw new Error(`API error ${res.status}: ${errBody}`);
        }

        const data = await res.json();
        if (data.choices && data.choices[0]?.message?.content) {
            return data.choices[0].message.content;
        }
        throw new Error("Unexpected API response format: " + JSON.stringify(data).slice(0, 300));
    }

    function buildPrompt(placeholderList, content) {
        const systemPrompt = `You are a document-filling assistant. You will receive user content and a list of template placeholder names. Your job is to extract the relevant information from the content and return a JSON object mapping each placeholder to its value.

RULES:
- Return ONLY a valid JSON object. No explanation, no markdown, no extra text.
- Each key must be one of the provided placeholder names.
- If a placeholder maps to a date, format as "Month Day, Year".
- If no content matches a placeholder, use "N/A".`;

        const userMessage = `PLACEHOLDERS: ${placeholderList.join(", ")}

CONTENT:
${content}

Return the JSON mapping now:`;

        return { systemPrompt, userMessage };
    }

    function parseAIResponse(raw) {
        let text = raw.trim();
        // Strip markdown code fences if present
        text = text.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/i, "");
        // Try to extract a JSON object
        const start = text.indexOf("{");
        const end   = text.lastIndexOf("}");
        if (start === -1 || end === -1) throw new Error("No JSON found in AI response: " + text.slice(0, 200));
        const jsonStr = text.slice(start, end + 1);
        try {
            return JSON.parse(jsonStr);
        } catch {
            // Try to fix common issues: trailing commas
            const fixed = jsonStr.replace(/,\s*}/g, "}").replace(/,\s*]/g, "]");
            return JSON.parse(fixed);
        }
    }

    // ── Process Button ────────────────────────────────────────
    processBtn.addEventListener("click", async () => {
        const content = userContentArea.value.trim();
        if (!content || !templateHtml) return;

        const btnText    = $(".btn-text");
        const btnLoading = $(".btn-loading");
        hide(btnText);
        show(btnLoading);
        processBtn.disabled = true;
        hide(outputSection);

        try {
            // Step 1: parse
            setProgress("parse");
            const pList = placeholders.length ? placeholders : guessSectionPlaceholders(templateHtml);
            await sleep(300);

            // Step 2: AI
            setProgress("ai");
            const { systemPrompt, userMessage } = buildPrompt(pList, content);
            const aiRaw = await callAI(systemPrompt, userMessage);
            const mappings = parseAIResponse(aiRaw);
            currentMappings = {};
            for (const p of pList) {
                currentMappings[p] = mappings[p] || mappings[p.toLowerCase()] || "";
            }

            // Step 3: fill
            setProgress("fill");
            await sleep(200);

            // Step 4: render
            setProgress("render");
            await renderPreview();
            progressDone();

            show(outputSection);
            outputSection.scrollIntoView({ behavior: "smooth", block: "start" });
        } catch (err) {
            alert("Error: " + err.message);
            hide(progressBar);
        } finally {
            show(btnText);
            hide(btnLoading);
            updateButtons();
        }
    });

    function guessSectionPlaceholders(html) {
        // Fallback: create placeholders for headings if none found
        const div = document.createElement("div");
        div.innerHTML = html;
        const headings = div.querySelectorAll("h1, h2, h3");
        if (headings.length) {
            return [...headings].map(h =>
                h.textContent.trim().toUpperCase().replace(/[^A-Z0-9]+/g, "_")
            );
        }
        return ["TITLE", "CONTENT"];
    }

    function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

    // ── Core: build a filled docx ZIP from originalDocxBuffer + currentMappings ──
    async function buildFilledDocxZip() {
        const zip = await JSZip.loadAsync(originalDocxBuffer);
        const xmlFiles = Object.keys(zip.files).filter(f =>
            f.endsWith(".xml") && !f.includes("_rels")
        );
        for (const fname of xmlFiles) {
            let xml = await zip.file(fname).async("string");
            let changed = false;
            xml = normalizeDocxParagraphs(xml);
            for (const [key, value] of Object.entries(currentMappings)) {
                const token = "{{" + key + "}}";
                if (!xml.includes(token)) continue;
                const lines = value.split("\n").filter(l => l.trim() !== "");
                if (lines.length <= 1) {
                    xml = xml.split(token).join(escapeXml(value));
                    changed = true;
                } else {
                    const pParts = xml.split(/(<w:p[\s>][\s\S]*?<\/w:p>)/);
                    const rebuilt = [];
                    for (const part of pParts) {
                        if (part.includes(token) && part.startsWith("<w:p")) {
                            const pPrMatch = part.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
                            const pPr = pPrMatch ? pPrMatch[0] : "";
                            const rPrMatch = part.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
                            const rPr = rPrMatch ? rPrMatch[0] : "";
                            const newParas = [];
                            for (const line of lines) {
                                const olMatch = line.match(/^\s*(\d+)[.)]\s+(.*)/);
                                const ulMatch = line.match(/^\s*[-\u2022*]\s+(.*)/);
                                const text = olMatch ? olMatch[2] : (ulMatch ? ulMatch[1] : line);
                                const isListItem = olMatch || ulMatch;
                                let itemPPr = pPr;
                                if (isListItem) {
                                    const indent = `<w:ind w:left="720" w:hanging="360"/>`;
                                    if (itemPPr && !itemPPr.includes("<w:ind ")) {
                                        itemPPr = itemPPr.replace("</w:pPr>", indent + "</w:pPr>");
                                    } else if (!itemPPr) {
                                        itemPPr = `<w:pPr>${indent}</w:pPr>`;
                                    }
                                    const prefix = olMatch ? `${olMatch[1]}. ` : "\u2022 ";
                                    newParas.push(`<w:p>${itemPPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(prefix + text)}</w:t></w:r></w:p>`);
                                } else {
                                    newParas.push(`<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`);
                                }
                            }
                            rebuilt.push(newParas.join("\n"));
                        } else {
                            rebuilt.push(part);
                        }
                    }
                    xml = rebuilt.join("");
                    changed = true;
                }
            }
            if (changed) zip.file(fname, xml);
        }
        return zip;
    }

    // Convert a filled docx ZIP to HTML via Mammoth
    async function filledDocxToHtml() {
        const zip = await buildFilledDocxZip();
        const buf = await zip.generateAsync({ type: "arraybuffer" });
        const result = await mammoth.convertToHtml({ arrayBuffer: buf });
        return result.value;
    }

    // ── Render & Edit ─────────────────────────────────────────
    async function renderPreview() {
        if (originalDocxBuffer) {
            // Best path: render from filled docx XML via Mammoth
            documentPreview.innerHTML = await filledDocxToHtml();
        } else {
            // Fallback for demo template or other cases
            documentPreview.innerHTML = fillTemplate(templateHtml, currentMappings);
        }
        buildEditPanel();
    }

    function buildEditPanel() {
        mappingsEditor.innerHTML = "";
        for (const [key, val] of Object.entries(currentMappings)) {
            const row = document.createElement("div");
            row.className = "mapping-row";
            row.innerHTML = `<label>{{${key}}}</label>
                <textarea data-key="${key}">${val}</textarea>`;
            mappingsEditor.appendChild(row);
        }
    }

    editToggleBtn.addEventListener("click", () => {
        editPanel.classList.toggle("hidden");
    });

    applyEditsBtn.addEventListener("click", async () => {
        mappingsEditor.querySelectorAll("textarea").forEach(ta => {
            currentMappings[ta.dataset.key] = ta.value;
        });
        await renderPreview();
    });

    // ── Manual Fill ───────────────────────────────────────────
    manualBtn.addEventListener("click", () => {
        manualFields.innerHTML = "";
        const pList = placeholders.length ? placeholders : ["TITLE", "CONTENT"];
        for (const p of pList) {
            const row = document.createElement("div");
            row.className = "mapping-row";
            row.innerHTML = `<label>{{${p}}}</label>
                <textarea data-key="${p}" rows="2" placeholder="Enter value for ${p}"></textarea>`;
            manualFields.appendChild(row);
        }
        show(manualModal);
    });

    manualCancelBtn.addEventListener("click", () => hide(manualModal));

    manualApplyBtn.addEventListener("click", async () => {
        currentMappings = {};
        manualFields.querySelectorAll("textarea").forEach(ta => {
            currentMappings[ta.dataset.key] = ta.value;
        });
        await renderPreview();
        hide(manualModal);
        show(outputSection);
        outputSection.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    // ── Word (.docx) Download ───────────────────────────────
    // Modifies the ORIGINAL uploaded .docx: replaces placeholders in ALL XML
    // files (body, headers, footers). Handles Word's run-splitting.
    // Multi-line values become separate <w:p> paragraphs.
    // Numbered lines (1. 2. 3.) get proper Word list indentation.
    downloadDocxBtn.addEventListener("click", async () => {
        if (!originalDocxBuffer) {
            alert("No template loaded. Please upload a .docx first.");
            return;
        }

        const zip = await JSZip.loadAsync(originalDocxBuffer);
        const xmlFiles = Object.keys(zip.files).filter(f =>
            f.endsWith(".xml") && !f.includes("_rels")
        );

        for (const fname of xmlFiles) {
            let xml = await zip.file(fname).async("string");
            let changed = false;

            // Step 1: Normalize fragmented runs within paragraphs.
            // Word splits {{PLACEHOLDER}} across multiple <w:r><w:t> elements.
            // We merge <w:t> text per-paragraph so replacements can match.
            xml = normalizeDocxParagraphs(xml);

            // Step 2: Replace each placeholder
            for (const [key, value] of Object.entries(currentMappings)) {
                const token = "{{" + key + "}}";
                if (!xml.includes(token)) continue;

                // For multi-line / list values: need to expand a single <w:p> into multiple
                const safeVal = escapeXml(value);
                const lines = value.split("\n").filter(l => l.trim() !== "");

                if (lines.length <= 1) {
                    // Simple single-line replacement
                    xml = xml.split(token).join(safeVal);
                    changed = true;
                } else {
                    // Multi-line: find each <w:p> containing the token and expand it
                    const pParts = xml.split(/(<w:p[\s>][\s\S]*?<\/w:p>)/);
                    const rebuilt = [];
                    for (const part of pParts) {
                        if (part.includes(token) && part.startsWith("<w:p")) {
                            // Extract paragraph & run properties for reuse
                            const pPrMatch = part.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
                            const pPr = pPrMatch ? pPrMatch[0] : "";
                            const rPrMatch = part.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
                            const rPr = rPrMatch ? rPrMatch[0] : "";

                            // Build replacement paragraphs
                            const newParas = [];
                            let listNum = 0;
                            for (const line of lines) {
                                const olMatch = line.match(/^\s*(\d+)[.)]\s+(.*)/);
                                const ulMatch = line.match(/^\s*[-•*]\s+(.*)/);
                                const text = olMatch ? olMatch[2] : (ulMatch ? ulMatch[1] : line);
                                const isListItem = olMatch || ulMatch;

                                let itemPPr = pPr;
                                if (isListItem) {
                                    listNum++;
                                    // Add hanging indent for list items
                                    const indent = `<w:ind w:left="720" w:hanging="360"/>`;
                                    if (itemPPr) {
                                        if (itemPPr.includes("<w:ind ")) {
                                            itemPPr = itemPPr.replace(/<w:ind [^/]*\/>/, indent);
                                        } else {
                                            itemPPr = itemPPr.replace("</w:pPr>", indent + "</w:pPr>");
                                        }
                                    } else {
                                        itemPPr = `<w:pPr>${indent}</w:pPr>`;
                                    }
                                    // Prefix with number or bullet
                                    const prefix = olMatch ? `${olMatch[1]}. ` : "• ";
                                    newParas.push(
                                        `<w:p>${itemPPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(prefix + text)}</w:t></w:r></w:p>`
                                    );
                                } else {
                                    newParas.push(
                                        `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`
                                    );
                                }
                            }
                            rebuilt.push(newParas.join("\n"));
                        } else {
                            rebuilt.push(part);
                        }
                    }
                    xml = rebuilt.join("");
                    changed = true;
                }
            }

            if (changed) {
                zip.file(fname, xml);
            }
        }

        const blob = await zip.generateAsync({
            type: "blob",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "docufill-output.docx";
        a.click();
        URL.revokeObjectURL(a.href);
    });

    // Merge fragmented <w:t> runs within each <w:p> so that
    // {{PLACEHOLDER}} tokens become contiguous in the XML text.
    function normalizeDocxParagraphs(xml) {
        return xml.replace(/<w:p[\s>][\s\S]*?<\/w:p>/g, (pBlock) => {
            // Collect all <w:t> text
            const texts = [];
            const tRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
            let m;
            while ((m = tRe.exec(pBlock)) !== null) texts.push(m[1]);
            const fullText = texts.join("");

            // Only rewrite if there's a placeholder in the joined text
            // that wasn't in a single run (i.e., it was fragmented)
            if (!/\{\{[A-Za-z0-9_ -]+\}\}/.test(fullText)) return pBlock;
            // Check if it's already intact in the original
            if (/\{\{[A-Za-z0-9_ -]+\}\}/.test(pBlock.replace(/<[^>]+>/g, ""))) {
                // Might already be fine, but let's check the raw XML text nodes
                const singleRunHas = texts.some(t => /\{\{[A-Za-z0-9_ -]+\}\}/.test(t));
                if (singleRunHas) return pBlock;
            }

            // Rewrite: preserve <w:pPr> and first <w:rPr>, replace all runs with one merged run
            const pPrMatch = pBlock.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
            const pPr = pPrMatch ? pPrMatch[0] : "";
            const rPrMatch = pBlock.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
            const rPr = rPrMatch ? rPrMatch[0] : "";

            return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${fullText}</w:t></w:r></w:p>`;
        });
    }

    function escapeRegex(str) {
        return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }

    // ── Demo .docx Generator ──────────────────────────────────
    async function generateDemoDocx() {
        const zip = new JSZip();

        // [Content_Types].xml
        zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

        // _rels/.rels
        zip.folder("_rels").file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

        // word/_rels/document.xml.rels
        zip.folder("word").folder("_rels").file("document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`);

        // word/document.xml — the actual template
        const body = DEMO_TEMPLATE_PARAGRAPHS.map(p => {
            if (p.heading) {
                const level = p.level || 1;
                return `<w:p>
  <w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>
  <w:r><w:t>${escapeXml(p.text)}</w:t></w:r>
</w:p>`;
            }
            // bold line
            if (p.bold) {
                return `<w:p>
  <w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(p.text)}</w:t></w:r>
</w:p>`;
            }
            return `<w:p>
  <w:r><w:t xml:space="preserve">${escapeXml(p.text)}</w:t></w:r>
</w:p>`;
        }).join("\n");

        zip.folder("word").file("document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
${body}
  </w:body>
</w:document>`);

        return await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    }

    function escapeXml(s) {
        return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
    }

    // ── Demo Template Content ─────────────────────────────────
    const DEMO_TEMPLATE_PARAGRAPHS = [
        { heading: true, level: 1, text: "{{PROJECT_NAME}} — Project Report" },
        { bold: true, text: "Prepared by: {{AUTHOR}}" },
        { bold: true, text: "Department: {{DEPARTMENT}}" },
        { bold: true, text: "Date: {{DATE}}" },
        { text: "" },
        { heading: true, level: 2, text: "1. Executive Summary" },
        { text: "{{SUMMARY}}" },
        { text: "" },
        { heading: true, level: 2, text: "2. Key Findings" },
        { text: "{{KEY_FINDINGS}}" },
        { text: "" },
        { heading: true, level: 2, text: "3. Recommendations" },
        { text: "{{RECOMMENDATIONS}}" },
        { text: "" },
        { heading: true, level: 2, text: "4. Conclusion" },
        { text: "{{CONCLUSION}}" },
        { text: "" },
        { text: "— End of Report —" }
    ];

    const DEMO_CONTENT = `Project Name: Smart Dashboard v2.0

This report was written by Jane Smith from the Analytics & Insights Department on March 15, 2026.

Executive Summary:
The Smart Dashboard v2.0 project was initiated to replace the legacy reporting system used across multiple business units. Over a 6-month development cycle, the team delivered a modern, real-time analytics platform that consolidates data from 12 internal sources into a single unified view. The project was completed on time and 8% under budget.

Key Findings:
After 3 months of pilot testing with 150 users across Sales, Marketing, and Operations, the following results were observed:
- Decision-making speed improved by 35% compared to the old system.
- Report generation time dropped from an average of 4 hours to under 15 minutes.
- User satisfaction scores averaged 4.6 out of 5, with particularly high marks for the customizable widgets and mobile responsiveness.
- Data accuracy improved by 22% due to automated validation pipelines.
- Three critical data silos were eliminated, enabling cross-department analysis for the first time.

Recommendations:
Based on the pilot results, we recommend the following next steps:
1. Roll out Smart Dashboard v2.0 to all remaining departments (Finance, HR, Legal) by Q3 2026.
2. Invest in advanced predictive analytics modules, leveraging the existing data pipelines.
3. Establish a dedicated Dashboard Support Team of 2-3 analysts to handle onboarding, training, and customization requests.
4. Schedule quarterly reviews to assess usage metrics and identify opportunities for new widget development.
5. Explore integration with external partner data sources to enrich market analysis capabilities.

Conclusion:
Smart Dashboard v2.0 has proven to be a significant upgrade over the legacy system, delivering measurable improvements in efficiency, accuracy, and user satisfaction. With a phased company-wide rollout and continued investment in advanced features, the platform is well-positioned to become the organization's central intelligence hub for data-driven decision-making.`;

    // ── Init ──────────────────────────────────────────────────
    updateButtons();

})();
