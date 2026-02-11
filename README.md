# Lark Base Printing & Typesetting Plugin

This is a sidebar plugin for Lark (Feishu) Base that allows you to:
1. Upload a Word (.docx) template with placeholders (e.g., `{{name}}`).
2. Select a record in the Base.
3. Automatically replace placeholders with data from the record.
4. Export the result as a PDF and save it back to an Attachment field in the Base.

## Development

1. Install dependencies:
   ```bash
   npm install
   ```
2. Start local server:
   ```bash
   npm run dev
   ```
3. Use the local server URL (e.g., `http://localhost:5173/`) in Lark Base Plugin setup.

## Deployment (GitHub Pages)

This project is configured to deploy to GitHub Pages automatically using GitHub Actions.

1. Push this code to a GitHub repository.
2. Go to **Settings** -> **Pages**.
3. Under **Build and deployment**, select **GitHub Actions** as the source.
4. The Action will run on push to `main` branch and deploy to the `gh-pages` environment.
5. Once deployed, you will get a URL (e.g., `https://yourname.github.io/repo-name/`).
6. Use this URL to install the plugin in Lark.

## Features

- **Word Templating**: Uses `docxtemplater` to replace `{{keys}}` with record values.
- **PDF Export**: Renders the DOCX to HTML client-side and converts to PDF (Note: layout may vary slightly from Word).
- **Auto-Save**: Uploads the generated PDF back to the first Attachment field found in the table.

## Notes

- Ensure your Base table has an **Attachment** field if you want the "Save to Base" feature to work.
- Complex Word formatting might not be perfectly preserved in the PDF preview/export due to client-side limitations.
