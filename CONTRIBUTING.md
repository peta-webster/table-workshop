# Contributing / 参与贡献

欢迎提交小而明确的改进：可复现的 Office 兼容性案例、错误修复、文档和示例。

## 开发

需要 Node.js 22+。运行 `npm ci` 安装测试依赖，`npm start` 启动应用，`npm test` 检查数据与文件结构。`npm run build` 输出静态文件到 `dist/`。

请修改 `src/`，不要手动修改生成的 `dist/`。保持每个 PR 聚焦一个问题，并写明触发条件、修改后的表现和验证方法。兼容性修改请加入最小回归案例，并说明是否在真实 Office 中验证；尚未实测的平台请明确标注。

## 报告问题

使用 Issue 模板，附上最小可复现表格、操作系统、浏览器、Office 名称/版本、复制还是下载、粘贴选项、预期和实际结果。截图很有帮助，但可复现输入更重要。请使用合成或脱敏数据。

## 依赖

浏览器依赖固定在 `src/vendor/`。升级时同时更新版本、完整许可证和 `integrity.json`，验证受影响的导出流程。不要加入不必要的网络请求或账户系统。

## English

Small, focused contributions are welcome. Use Node.js 22+, run `npm ci`, `npm test`, and `npm run build`. Edit `src/`, not generated `dist/`. Describe the reproduction, resulting behavior, and validation in your PR. Include a minimal regression case for compatibility fixes and distinguish structural tests from actual Office verification.

For bug reports, include sanitized input, OS/browser/Office versions, copy or download path, paste option, and expected/actual behavior. Dependency upgrades must update the bundled files, notices, versions and integrity hashes together.

By contributing, you agree to license your contributions under the project's MIT license.
