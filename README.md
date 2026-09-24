# 表格工坊 · Table Workshop

把 AI / Markdown 表格转换成适合 Word、Excel 和 PowerPoint 的可编辑表格。

[English](README.en.md) · [下载版本](https://github.com/peta-webster/table-workshop/releases) · [兼容性记录](docs/compatibility.md)

![表格工坊界面](docs/screenshot.png)

## 能做什么

粘贴一张表格 → 选择目标软件与样式 → 复制，或下载 Office 文件。

| 目标 | 适合的任务 | 使用方式 |
| --- | --- | --- |
| Word | 报告、论文、三线表 | 复制后选择「保留源格式」；或下载 `.docx` |
| Excel | 把数据放进已有工作表 | 默认「无样式（仅数据）」；粘贴后选择「匹配目标格式」；或下载 `.xlsx` |
| PowerPoint | 汇报中的可编辑表格 | 下载 `.pptx`，长表自动拆页并重复表头 |

- 支持 Markdown、网页 HTML 表格和简单制表符文本。
- 保留中文、单元格换行、空值、编号前导零与长数字。
- Word 复制宽度随正文区域调整，文字颜色和段落间距明确设置。
- Excel 区分数值与文本；公式样式文本不会作为公式执行。
- 所有表格内容在浏览器本地处理，无账号、AI API、后端接口、分析统计或运行时 CDN 请求。访问在线页面时，静态托管服务仍会收到普通网页请求。

### Excel「无样式」的含义

只保留数据与必要的数字显示格式。下载文件从 A1 开始，不插入额外标题、配色、边框、固定行高列宽、合并单元格、筛选或冻结行。预览中的辅助网格不导出。

复制时仍携带单元格类型，避免 `00123` 和长编号被改写。粘贴到已有 Excel 后选择「匹配目标格式」，外观和数字显示沿用目标单元格：例如数值 `0.985` 在「常规」格式下显示为 `0.985`，在百分比格式下显示为百分比。需要保留原来的小数位或百分比显示时，下载 `.xlsx`。

## 本地使用

安装 [Node.js](https://nodejs.org/) 22 或更高版本，然后：

```sh
git clone https://github.com/peta-webster/table-workshop.git
cd table-workshop
npm start
```

打开 http://127.0.0.1:4173/ ，保持终端运行。启动和构建使用 Node 内置模块，无需先安装 npm 依赖。源码修改后重启即可更新页面。

也可以下载 Releases 中的 `table-workshop-v0.1.0-web.zip`，解压后运行 `node scripts/serve.mjs`。直接双击 HTML 文件不支持模块加载。

### 开发与验证

```sh
npm ci
npm test
npm run build
npm run package
```

测试检查输入结构、数据类型、剪贴板 HTML，以及生成的 DOCX/XLSX/PPTX 内部结构。真实 Office 的粘贴与排版验证另见[兼容性记录](docs/compatibility.md)，浏览器预览不代表 Office 的逐像素渲染。

仓库示例、测试输入和展示截图均使用虚构数据。

## 当前边界

- 一次处理一张矩形表格，最多 30 列、10,000 个数据单元格；PPT 最多 10 列。
- 暂不支持合并单元格。列数不一致会提示，超长 PPT 单行会要求拆分。
- 单元格内容统一为文本；链接保留文字与 HTTP(S) 地址，图片仅保留替代文本。复杂富文本与数学公式保持文本表达。
- 制表符输入适用于每行一个记录的简单文本，不支持完整 CSV/带引号 TSV 文件导入。
- 文件导出比跨软件复制更可控。Windows Office、WPS 和其他浏览器的粘贴效果仍待补充验证。
- 界面目前为简体中文，文档提供中文和英文。

## 项目结构

```text
src/core/       表格解析、类型保护、列宽与分页
src/exporters/  Office 文件与剪贴板输出
src/app.js      页面交互与预览
src/vendor/     固定版本浏览器依赖、许可证和校验值
scripts/        本地服务、构建与打包
tests/          回归检查
examples/       可以直接粘贴的示例
docs/           兼容性与发布说明
```

部署：运行 `npm run build`，把 `dist/` 放到静态 HTTPS 托管服务。资源使用相对路径，支持子目录部署。剪贴板权限由浏览器管理；复制不可用时可下载文件。

## 参与

欢迎贡献可复现的兼容性案例，尤其是 Windows Office、WPS、不同浏览器和中文长表。反馈时附上最小示例、软件版本与复制/下载路径，使用合成或脱敏数据即可。见 [贡献指南](CONTRIBUTING.md)、[行为准则](CODE_OF_CONDUCT.md) 和 [更新记录](CHANGELOG.md)。

后续优先根据真实问题改善兼容性，再评估浏览器扩展与可复用接口。

## 许可证

项目代码采用 [MIT](LICENSE)。第三方库保留各自许可证，详见 [第三方声明](THIRD_PARTY_NOTICES.md)。
