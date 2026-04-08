# DocuFill AI

Smart AI-powered template filler — upload a Word template with `{{PLACEHOLDERS}}`, provide your content, and let AI fill it in automatically.

## Features

- **Template Upload** — Drag & drop or browse for `.docx` Word templates
- **Placeholder Detection** — Automatically finds `{{PLACEHOLDER}}` tokens in document body, headers, and footers
- **AI-Powered Filling** — Analyzes your content and maps it to template placeholders using Hugging Face Inference API
- **Manual Fill** — Option to manually map placeholders without AI
- **Word Download** — Downloads a filled `.docx` that preserves original template formatting, styles, headers, and footers
- **Live Preview** — Approximate preview of the filled document in the browser
- **Demo Template** — Built-in sample template for quick testing

## How It Works

1. **Setup** — Enter your Hugging Face API key (saved in browser localStorage)
2. **Upload Template** — Upload a `.docx` file containing `{{PLACEHOLDER}}` tokens (or use the demo template)
3. **Provide Content** — Paste the source content you want inserted into the template
4. **Analyze & Fill** — AI reads your content and maps it to each placeholder
5. **Download** — Download the filled Word document

## Getting Started

### Online (GitHub Pages)

Visit the live app and enter your Hugging Face API key in the Setup section.
https://rakspa-ts.github.io/docufill-ai/

### Local Development

1. Clone the repo:
   ```bash
   git clone https://github.com/rakspa-ts/docufill-ai.git
   cd docufill-ai
   ```

2. (Optional) Create a `.env` file with your API key:
   ```
   HF_API_KEY=hf_your_token_here
   ```

3. Start a local server:
   ```bash
   python -m http.server 8080
   ```

4. Open http://localhost:8080

### Hugging Face API Key

You need a **free** Hugging Face token with the **"Make calls to Inference Providers"** permission.

Create one here: [Hugging Face Token Settings](https://huggingface.co/settings/tokens/new?ownUserPermissions=inference.serverless.write&tokenType=fineGrained)

## Tech Stack

- **Frontend** — Pure HTML, CSS, JavaScript (no framework)
- **Word Parsing** — [Mammoth.js](https://github.com/mwilliamson/mammoth.js) 1.6.0
- **ZIP Handling** — [JSZip](https://stuk.github.io/jszip/) 3.10.1
- **AI** — [Hugging Face Inference API](https://huggingface.co/docs/api-inference/) (OpenAI-compatible chat completions)
- **Models** — Meta Llama 3.1 8B Instruct (default), Qwen 2.5 72B Instruct, DeepSeek R1

## Template Format

Templates are standard `.docx` files with placeholders in double curly braces:

```
Dear {{RECIPIENT_NAME}},

This letter confirms your role as {{JOB_TITLE}} starting {{START_DATE}}.
```

Placeholders can appear in the document body, headers, and footers.

## Project Structure

```
├── index.html          # Main app page
├── css/styles.css      # Styling
├── js/app.js           # Application logic
├── .env.example        # API key template
└── README.md
```

## License

MIT
