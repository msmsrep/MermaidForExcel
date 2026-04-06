# Mermaid for Excel

[English](./README.md) | [日本語](./README.ja.md)

An Office Add-in that creates diagrams using Mermaid syntax and inserts them as images into Excel sheets.

Quickly draw flowcharts, sequence diagrams, ER diagrams, and more using text-based notation, then insert them as PNG/JPEG images into the active sheet. All processing happens locally — your input is never sent to external servers.

### Key Features

- **Real-time preview** — Enter Mermaid syntax and click Render to instantly preview the diagram in the task pane
- **Insert into Excel** — Insert the previewed image into the active sheet with one click
- **PNG / JPEG output** — Choose your output format and download the image directly
- **Wide diagram support** — All diagram types supported by [Mermaid](https://github.com/mermaid-js/mermaid) are available, including flowcharts, sequence diagrams, class diagrams, ER diagrams, and Gantt charts
- **Privacy-friendly** — All processing is completed within the task pane; no data is sent to external servers

## Project Structure

```
MermaidForExcel/
  build.js                 ← esbuild build script (serve / watch / build)
  manifest.xml             ← Office Add-in manifest
  package.json
  tsconfig.json
  assets/                  ← Icon directory (placeholder icons are auto-generated on first build)
  src/taskpane/
    taskpane.html
    taskpane.ts
  dist/                    ← Build output (not tracked by git)
```

## Setup

```bash
cd MermaidForExcel
npm install

# Install development HTTPS certificate (first time only, requires admin privileges)
npx office-addin-dev-certs install
```

## Starting the Development Server

```bash
npm start
# → https://localhost:3000/taskpane.html
```

## Registering the Add-in in Excel

1. Open Excel
2. **[Insert]** → **[Add-ins]** → **[My Add-ins]** → **[Upload My Add-in]**
3. Select `dist/manifest.xml`

An **[Open Mermaid]** button will be added to the Home tab.

## Scripts

| Command | Description |
|---|---|
| `npm start` | Build + start HTTPS server (port 3000) |
| `npm run build` | One-time build to `dist/` |
| `npm run watch` | Watch for file changes and auto-rebuild |

## Replacing Icons

Replace `assets/icon-{16,32,80}.png` with your own PNG images.  
1×1px placeholder icons are auto-generated on the first build.

## License

This project is licensed under the [MIT License](LICENSE.txt).

### Third-party licenses

This software uses [mermaid](https://github.com/mermaid-js/mermaid) (MIT License).

## Privacy Policy

Last updated: April 6, 2026

### Data Collection

Mermaid for Excel does **not** collect any personal information or user data.

### How It Works

- Mermaid syntax text entered by the user is processed **entirely within the task pane** in Excel
- Diagram rendering is completed locally and no data is sent to external servers
- Inserting images into Excel sheets is also performed locally via the Microsoft Office JavaScript API

### Access to External Services

This add-in loads its application files (HTML/JavaScript) from a hosting server.  
No user data is included in this communication.

### Cookies and Tracking

This add-in does not use cookies, local storage, or any tracking technologies.

### Contact

For privacy-related questions, please open a [GitHub Issue](https://github.com/msmsrep/MermaidForExcel/issues).
