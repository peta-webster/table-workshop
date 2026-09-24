# Table Workshop

Turn AI and Markdown tables into editable tables for Word, Excel and PowerPoint. Table content stays in your browser.

[中文](README.md) · [Releases](https://github.com/peta-webster/table-workshop/releases) · [Compatibility](docs/compatibility.md)

![Table Workshop](docs/screenshot.png)

## Workflow

Paste one table, choose the target app and style, then copy or download.

- **Word:** editable DOCX, academic three-line tables, clipboard tables that follow the document text width. Choose **Keep Source Formatting** after pasting.
- **Excel:** defaults to **Data only**. No decorative fonts, fills, borders, extra title, or fixed row/column dimensions. Choose **Match Destination Formatting** when pasting into an existing sheet.
- **PowerPoint:** native editable tables, automatic pagination and repeated headers. Download PPTX for the most predictable result.

Markdown, HTML tables and simple tab-separated text are supported. Chinese text, line breaks, blanks, leading-zero identifiers and long numbers are preserved. Formula-looking input stays text. Percentages and ordinary numbers remain numeric where clearly identified.

**Data only still preserves types.** In XLSX, necessary number formats preserve decimal places and percentage display. When matching destination formatting during paste, the destination controls display: a percentage value of `0.985` appears as `0.985` in a General cell and as a percentage in a percentage-formatted cell.

## Run locally

Install Node.js 22 or newer:

```sh
git clone https://github.com/peta-webster/table-workshop.git
cd table-workshop
npm start
```

Open http://127.0.0.1:4173/ and keep the terminal running. Starting and building use built-in Node modules and do not require an npm install. Restart after editing source files.

Alternatively, download the web ZIP from Releases, extract it and run `node scripts/serve.mjs`. Opening the HTML directly through `file://` does not support module loading.

## Develop

```sh
npm ci
npm test
npm run build
npm run package
```

Tests verify parsing, cell types, clipboard HTML, and the internal structures of exported DOCX/XLSX/PPTX files. These checks do not replace actual Office paste/layout testing. The compatibility document distinguishes tested workflows from pending coverage.

All public examples, test input and screenshots use synthetic data.

Deploy `dist/` to a static HTTPS host after building. Relative asset URLs support subdirectory hosting. The app has no conversion backend, AI API, analytics or runtime CDN requests. The static hosting provider still receives normal page requests. Clipboard access depends on browser permissions; file downloads are available when copying is unavailable.

## Scope

- One rectangular table per input; up to 30 columns and 10,000 data cells. PPT supports up to 10 columns.
- Merged cells and inconsistent row widths are rejected. Oversized single PPT rows require splitting.
- Cell formatting is normalized. Links retain their text and HTTP(S) URL; images retain alt text only. Rich text and mathematical formulas are not preserved as native formatting/equations.
- Tab-separated input supports simple one-record-per-line text, not a full quoted TSV/CSV importer.
- The interface is currently Simplified Chinese. Documentation is available in English and Chinese.
- macOS Office workflows have been tested. Windows Office, WPS and other browser combinations need community validation. Browser previews are not pixel-perfect Office renders.

See [examples](examples), [contributing](CONTRIBUTING.md), [code of conduct](CODE_OF_CONDUCT.md), and [changelog](CHANGELOG.md). Reproducible compatibility reports are particularly welcome.

## License

Project code: [MIT](LICENSE). Bundled libraries retain their licenses; see [third-party notices](THIRD_PARTY_NOTICES.md).
