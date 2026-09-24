# Third-party notices

The application bundles browser distributions so runtime conversion does not depend on a CDN. These files keep their upstream licenses and are not covered solely by this project's LICENSE.

| Package | Version | License notice | Upstream |
| --- | --- | --- | --- |
| docx | 9.6.1 | [MIT notice](src/vendor/docx-LICENSE) | https://github.com/dolanmiu/docx |
| ExcelJS | 4.4.0 | [MIT notice](src/vendor/exceljs-LICENSE) | https://github.com/exceljs/exceljs |
| Marked | 17.0.5 | [MIT and Markdown notices](src/vendor/marked-LICENSE) | https://github.com/markedjs/marked |
| PptxGenJS | 4.0.1 | [MIT notice](src/vendor/pptxgenjs-LICENSE) | https://github.com/gitbrent/PptxGenJS |

Exact bundle and license-file SHA-256 hashes are in `src/vendor/integrity.json`. Every build checks them. The web release includes these notices under `dist/vendor/`.

Development-only dependencies are pinned in `package.json` and `package-lock.json`; their upstream license files are distributed with their npm packages. They are not loaded by the application.

When upgrading a bundled library, obtain the browser distribution and complete notices from the corresponding upstream package, update the version and hashes together, and run the full test suite plus the affected Office checks. A hash verifies that a checked-in file has not changed; it is not a claim of a reproducible upstream build.
