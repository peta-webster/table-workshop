# Table Workshop

Turn AI and Markdown tables into editable tables for Word, Excel and PowerPoint. Table content stays in your local browser.

[中文](README.md) · [Releases](https://github.com/peta-webster/table-workshop/releases) · [Example inputs](examples/README.md) · [Compatibility notes](docs/compatibility.md)

## Quick start

Install [Node.js](https://nodejs.org/) 22 or newer, then run:

```sh
git clone https://github.com/peta-webster/table-workshop.git
cd table-workshop
npm start
```

Open the `http://127.0.0.1:4173/` address printed in the terminal and keep the server running. Startup does not require an npm install. You can also download the web ZIP from [Releases](https://github.com/peta-webster/table-workshop/releases), extract it, and run `node scripts/serve.mjs`.

## How to use it

1. Paste one Markdown table, HTML table from a web page, or simple tab-separated text.
2. Choose Word, Excel or PowerPoint and a style.
3. Copy the table or download an editable Office file.

![Table Workshop input and preview with synthetic office-supply data](docs/screenshots/overview.jpg)

### Three browser previews

These images were captured from the local app using synthetic examples. They show browser previews; the files may look different in Office. Click an image to see it full size.

| Word · academic three-line style | Excel · data only | PowerPoint · presentation style |
| --- | --- | --- |
| [![Browser preview of a Word three-line table](docs/screenshots/word-preview.png)](docs/screenshots/word-preview.png) | [![Browser preview of an Excel data-only table](docs/screenshots/excel-preview.png)](docs/screenshots/excel-preview.png) | [![Browser preview of a PowerPoint presentation table](docs/screenshots/powerpoint-preview.png)](docs/screenshots/powerpoint-preview.png) |
| [Project-status input](examples/project-status.md) | [Data-types input](examples/data-types.md) | [Project-status input](examples/project-status.md) |

## Output and paste guidance

| Target | What to use |
| --- | --- |
| Word | Download a native `.docx`, or copy and choose **Keep Source Formatting** in Word. Clipboard table widths follow the document text area. |
| Excel | **Data only** is the default. Identifiers, leading zeros, long numbers and formula-looking input stay text; clear numbers and percentages stay numeric. Choose **Match Destination Formatting** when pasting into an existing sheet. |
| PowerPoint | Download a native editable `.pptx`. Long tables split across slides with repeated headers; cross-app paste results depend on the Office version. |

In Excel, **Match Destination Formatting** also uses the destination's number format. Download `.xlsx` when the original percentage or decimal display must be preserved.

## Limits and validation

- One rectangular table at a time, up to 30 columns and 10,000 data cells; PowerPoint supports up to 10 columns.
- Merged cells, complex rich text and native equations are unsupported. Simple tab-separated input is not a full CSV or quoted-TSV importer.
- Automated tests check input, data types, clipboard output and Office file structure. See the [compatibility notes](docs/compatibility.md) for actual Office paste and layout coverage; Windows Office and WPS still need testing.

## Development and repository layout

```sh
npm ci
npm test
npm run build
npm run package
```

| Directory | Contents |
| --- | --- |
| `src/` | App, parsing, Office export and pinned browser libraries |
| `scripts/` | Local server, static build and release packaging |
| `tests/` | Regression tests |
| `examples/` | Copyable synthetic inputs and an example index |
| `docs/` | Compatibility notes and browser-preview screenshots |

The static build is written to `dist/` and can be deployed to an HTTPS static host. The app has no accounts, conversion backend, AI API, analytics or runtime CDN requests; the host still receives normal page requests.

See the [contributing guide](CONTRIBUTING.md), [code of conduct](CODE_OF_CONDUCT.md), and [changelog](CHANGELOG.md). Project code uses the [MIT license](LICENSE); bundled libraries retain their licenses in the [third-party notices](THIRD_PARTY_NOTICES.md).
