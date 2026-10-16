# E-Master Site 完整源码导出

这是当前线上 **第14版** E-Master 的现有项目文件，来自该发布版本记录的源码提交：

`1dc2ef2d67272f250b02721c28290551d23edc37`

线上地址：https://e-master.honest-olive-0596.chatgpt.site

原有代码、配置、依赖库和业务模板均按原字节复制。此次没有修改业务逻辑、重新开发、重构或重新发布站点。新增文件仅用于说明、导出核验和配置快照，详见 `SOURCE_EXPORT_MANIFEST.json`。

## 1. 技术栈与运行方式

本项目是 **HTML + CSS + 原生 JavaScript 的纯前端静态应用**。`dist/` 内的 JavaScript 文件就是当前实际使用的业务源码，不是重新生成的替代实现。

- Excel 读取和写入在用户浏览器中通过 JSZip 3.10.1、DOMParser 和 OOXML XML 处理完成。
- JSZip 已包含在 `dist/vendor/jszip.min.js`，保留当前原文件及其中打包的依赖，无需从CDN下载。
- 目标模板及结构校验文件通过同源 `fetch()` 从 `dist/templates/` 读取。
- 没有业务后端、服务端路由、数据库、Node服务或Python业务处理程序。
- 项目没有 `package.json`、npm锁文件或编译步骤，因此无需执行 `npm install` 或 `npm run build`。
- 页面和Excel处理不调用ChatGPT、OpenAI模型、MCP或Plugin。

## 2. 目录结构

| 路径 | 内容 |
| --- | --- |
| `dist/index.html` | 实际页面入口及脚本加载顺序 |
| `dist/style.css` | 当前页面样式 |
| `dist/app.js` | 上传、映射管理、进度、错误明细和下载交互 |
| `dist/engine.js` | 第一阶段标准化、日期/资产ABC校验、设备代码处理、默认设备编码映射 |
| `dist/assembly-rules.js` | 第二阶段字段别名映射、必填字段、关键表头和模板路径配置 |
| `dist/assembly.js` | 设备及技术参数校验、真实源行号、目标模板装配和回读保护 |
| `dist/vendor/jszip.min.js` | 现有JSZip依赖，包含打包的传递依赖 |
| `dist/templates/设备档案导入JLTY0714.xls` | 当前实际使用的原始固定目标模板，原字节保留 |
| `dist/templates/template-contract.json` | 当前模板Sheet、隐藏状态、表头及公式的结构校验基线 |
| `.openai/hosting.json` | 源码仓库中的Sites静态托管配置 |
| `dist/.openai/hosting.json` | 原发布包中的托管元数据，与根目录配置内容一致 |
| `docs/` | 仓库原有第二阶段及日期校验验收文档 |
| `tests/date-validation.cjs` | 仓库原有日期校验测试 |
| `export-info/default-device-code-mappings.json` | 从现有 `EM.MAP` 自动提取的11项默认映射快照，供查阅，不参与运行 |
| `export-info/DEPENDENCIES.json` | 导出新增的依赖清单，不参与运行 |
| `SITES_RUNTIME_DEPENDENCIES.md` | 平台能力及本地运行依赖的独立说明 |
| `SOURCE_EXPORT_MANIFEST.json` | 发布版本、来源提交、原始文件清单及逐文件SHA256 |
| `SHA256SUMS.txt` | 包内文件SHA256清单（不包含该清单自身） |

## 3. 本地启动（Windows）

前提：本机已安装Python 3，并有支持现代JavaScript和上述浏览器API的浏览器。

1. 解压 `E-Master-Site-Source.zip`。
2. 打开包含本README的项目目录。例如项目放在 `D:E-Master-Site-Source`，在PowerShell中运行：

```powershell
cd 'D:\E-Master-Site-Source'
py -3 -m http.server 8000 --bind 127.0.0.1 --directory dist
```

如果电脑没有 `py` 命令，但已配置 `python`，使用：

```powershell
python -m http.server 8000 --bind 127.0.0.1 --directory dist
```

3. 保持命令窗口运行，在浏览器打开：

```text
http://127.0.0.1:8000/
```

4. 上传Excel、补充所需映射并点击“开始处理”，即可下载最终文件。
5. 结束时回到命令窗口按 `Ctrl+C`。

Python在这里只提供静态文件，不执行Excel转换业务。不要直接双击 `dist/index.html` 使用 `file://` 打开，因为浏览器会限制模板的 `fetch()` 读取。

### macOS / Linux

在包含本README的目录运行：

```bash
python3 -m http.server 8000 --bind 127.0.0.1 --directory dist
```

然后打开同一个 `http://127.0.0.1:8000/`。

### 部署到其他静态服务器

把 `dist/` 中的全部文件部署到同一个静态站点，入口为 `index.html`，保持 `templates/`、`vendor/` 和各脚本的相对路径不变。无需反向代理业务API或连接数据库。

本机运行命令仅监听本机。若需要供局域网同事访问，可改为 `--bind 0.0.0.0`，再使用服务器局域网IP访问同一端口；操作系统防火墙需要允许该端口。原Sites的登录和访问控制不会随源码迁移，本地静态服务器也没有登录模块。

## 4. 编码映射和业务配置

**真正参与运行的默认编码表位于 `dist/engine.js` 的 `MAP`。** 当前有11项，导出快照见 `export-info/default-device-code-mappings.json`。快照和依赖清单都是导出附加文件，应用不会自动读取它们；编辑快照不会改变转换规则。

用户在页面中“手动添加”的映射另行保存在浏览器 `localStorage`，键为：

```text
e-master-custom-mappings-v1
```

这部分数据不在Sites源码仓库中，也不是服务器保存的配置。因此本ZIP不包含各浏览器中手动添加的映射（例如后来手动添加的安全阀映射）。本地首次打开后可以直接在原有页面中重新添加，操作方式不变。

如要迁移已有自定义映射：

1. 在原线上站点使用原来的浏览器和用户配置，按F12打开开发者工具。
2. 在“应用 / Application → 本地存储 / Local Storage”中选择原站点域名，复制上面键对应的完整JSON值。
3. 在本地 `http://127.0.0.1:8000/` 页面打开相同面板，在本地站点的Local Storage中创建同名键，粘贴该JSON值，然后刷新页面。
4. 展开“查看全部编码映射”确认迁移结果。也可直接在页面逐项重新添加。

默认表中的项不能被自定义项覆盖，这是原有行为。不同浏览器、域名、IP或端口的Local Storage互相独立；`localhost:8000` 与 `127.0.0.1:8000` 也是不同来源。

目标模板、字段映射和规则的位置：

| 内容 | 当前来源 |
| --- | --- |
| 固定Excel模板 | `dist/templates/设备档案导入JLTY0714.xls` |
| 模板保护基线 | `dist/templates/template-contract.json` |
| 源字段与目标字段的别名映射、必填配置 | `dist/assembly-rules.js` |
| 日期字段及100年年限配置 | `dist/engine.js` 中的 `DATE_FIELDS`、`MAX_DATE_AGE_YEARS` |

当前固定模板虽然扩展名为 `.xls`，其文件内容实际是OOXML ZIP。已按当前实际文件原样导出，未转换或重新保存。模板中所有24个Sheet、隐藏Sheet、业务内容、表头和公式都保留；请将模板与校验基线配套保留。

## 5. 当前完整处理流程

一次上传 `.xlsx` 文件后，程序先按表头执行现有第一阶段标准化，再校验设备编码、技术参数唯一性及关联关系，校验通过后装配固定模板并生成最终 `.xlsx`。无需下载和再次上传中间文件。

本次导出沿用当前全部业务行为，包括：

- 设备小类映射缺失时阻断；三列日期标准化和月份补01；仅年份、不存在的日期、未来年份、距当年超过100年的日期阻断。
- 资产ABC空值或 `/` 置空，单个A/B/C/a/b/c转大写，其他内容阻断。
- 设备代码按当前JSON规则处理；输出文件名使用原文件名、`_导入版_` 和北京时间月日。
- 同一设备同一参数重复即阻断，不自动去重；错误明细保留源Excel真实行号。
- 按字段名称映射，只写设备档案和技术参数；保留模板的全部Sheet及结构，并回读校验最终文件。

原项目没有服务器端保存的用户上传文件、历史处理结果或集中式自定义映射数据库，这些不属于此次源码内容。当前应用输入只支持 `.xlsx`；旧版二进制 `.xls` 源文件需先用Excel/WPS另存为 `.xlsx`。

## 6. ChatGPT Sites Runtime 依赖

独立清单见 `SITES_RUNTIME_DEPENDENCIES.md`。

核心Excel处理和页面功能不依赖ChatGPT Sites Runtime，不调用OpenAI模型或专有运行API。本地文件准备齐全后，不需要ChatGPT登录、API Key、VPN或外网连接来执行处理；浏览器只需能访问本机静态HTTP服务。

线上站点的托管、域名/HTTPS、平台登录与访问控制、源码版本及发布管理由Sites平台提供。这些平台内部服务不属于本项目可导出的源码，本地启动不会复制原站点的登录控制。

## 7. 测试与导出核验

原有日期测试可在另行安装支持 `node:test` 的Node.js后，于项目根目录执行：

```bash
node --test tests/date-validation.cjs
```

`docs/` 中的验收文档是原仓库已有记录；其中涉及的临时验收环境和用户样例不属于站点运行依赖。本次是源码导出，没有重新运行全部历史业务验收，也没有改动代码。

本次导出核验包括：当前线上第14版发布记录与Git提交一致；全部13个原仓库文件按Git blob内容原样导出；与保留的该版发布包比对了10个发布文件的内容；ZIP解压完整性、文件SHA256、脚本及模板相对路径完整性通过检查。原源码仓库保持未修改。

`.git/` 的本地历史、Git传输凭据、平台内部托管实现及用户浏览器存储不在本ZIP中；现有项目的全部代码、静态资源、配置、依赖bundle及固定业务模板均已包含。
"# E-Master-Offline" 
"# E-Master-Offline" 
"# E-Master-Offline" 
