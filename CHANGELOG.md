# Changelog

## 0.2.0 — 2026-10-06

- Add black-and-white print, light horizontal-line and compact-grid table styles across previews, clipboard output and Office exports.
- Add Modern Office (Microsoft YaHei / Arial) and Formal Document (SimSun / Times New Roman) font pairs, remembered per output target during the session. Font sizes and spacing remain preset.
- Preserve numeric values and text identifiers when applying font pairs in styled Excel exports; data-only Excel output continues to omit fonts and decoration.
- Document Word's Keep Source Formatting paste workflow and add online demo links to both READMEs.
- Expand automated regression coverage to 11 tests, including new styles and font pairs.

Validation boundary: automated tests inspect clipboard HTML and Office file structure. A user-provided Word screenshot showed better results with Keep Source Formatting than Merge Formatting; this is not comprehensive Windows Office or WPS validation. Fonts unavailable on the recipient's device may be substituted.

## 0.1.0 — 2026-09-24

First open-source release of Table Workshop.

- Markdown, HTML and simple TSV table input with explicit structure validation.
- Editable DOCX, XLSX and PPTX exports; Chinese text, line breaks and empty cells.
- Word clipboard widths follow the document text area; explicit header text formatting.
- Excel data-only output with identifier/type protection and destination-format paste guidance.
- PowerPoint long-table pagination and repeated headers.
- Local-only table processing, vendored browser dependencies, portable Node startup and static build.
- Regression tests and a record of actual macOS Office validation.

Known limitations: merged cells and complex rich text are unsupported; Windows Office and WPS paste/layout validation is pending. The interface is Simplified Chinese.
