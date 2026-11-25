<p align="center">
  <img src="./banner.svg" alt="RefLink Pro — 跨文献软件与无插件纯文本全能交叉引用及导出可点击 PDF 插件" width="100%">
</p>

# RefLink Pro 全能版 — 不依赖文献管理软件也能点击跳转 + 导出可点击 PDF（Word & WPS 全平台）

> **哪怕完全没用文献管理软件（手敲/知网复制的纯文本 `[1]`、`[2-5]`、上标 `¹⁻³`、作者年份 `(Smith, 2020)`），或者使用 EndNote、NoteExpress、Mendeley、Citavi、Zotero**，都能一键开启**软件内点击角标秒跳参考文献**，并在定稿副本一键导出**黑字、无下划线、保持原小上标不变的可点击 PDF**！

写毕业论文、医学/理工期刊初稿或合著统稿时，经常遇到这些棘手情况：
1. **手敲纯文本或已去域（Unlink）稿件**：从知网/PubMed 复制参考文献手敲角标 `[1]`、`[2-3]`，或多位作者合稿后引文域丢失，手工逐个插几百个交叉引用书签耗费数小时且极易错位；
2. **使用 EndNote / NoteExpress / Mendeley**：市面上多数无痕跳转工具只支持单一软件，换用 EndNote 或 NoteExpress 就无法无损跳转或导出干净的可点击 PDF；
3. **传统宏或自带超链接破坏排版**：一跑脚本正文角标全变**蓝色+下划线**，甚至把小上标撑大成正文字号，不符合学位论文盲审与期刊排版规范。

**RefLink Pro 全能版（v0.7.2）** 完整包含 **Plus 版 + 标准版全部功能**，并搭载全新**通用引文与参考文献智能识别引擎**：
- **纯文本与全文献软件通吃**：原生支持 **手敲纯文本引文**（`[1]`、`[1, 3-5]`、`［1］`、无括号真实上标 `¹⁻³`、中英文作者-年份 `(Smith, 2020)` / `张三等（2022）`），同时全面兼容 **EndNote、NoteExpress、Mendeley、Citavi、Word 内置引文与 Zotero**（包括同一篇论文里“文献插件 + 手敲纯文本”混合使用）！
- **学术级防误判保护**：自动识别文末「参考文献 / 参 考 文 献 / References」（自动跳过前文目录，在「致谢 / 附录 / 利益冲突」前自动截止），支持 Word/WPS 自动编号文献表，并**智能排除 `m²`、`kg/m²`、`cm³`、`10⁶`、`R²`、`χ²`、`(P<0.05)` 等单位与统计符号**，绝不乱跳或误加链接！
- **写作安全跳转 + 定稿无痕可点击 PDF 双模式**：日常改稿零蓝字、零下划线、不破坏活动引文域；定稿副本一键写入内部书签超链接并自动还原**纯黑字体、无下划线、原小上标大小**，导出 PDF 后外审专家与审稿人点击任意角标瞬间直达参考文献！

---

## 一、版本与方案全景对比

| 对比维度 | 手工插交叉引用 / 传统 VBA 宏 | [RefLink 标准版](https://github.com/huangshaoxiang123/RefLink) | [RefLink Plus 版](https://github.com/huangshaoxiang123/RefLink-Plus) | **RefLink Pro 全能版（本仓库旗舰）** |
| :--- | :--- | :--- | :--- | :--- |
| **支持平台** | 仅部分桌面 Word，常报错 | Word & WPS（Mac / Win） | Word & WPS（Mac / Win） | **Word & WPS（Mac / Win 全平台）** |
| **Zotero 活动引文域** | 易破坏 `Zotero Refresh` | 支持（保 Refresh） | 支持（保 Refresh） | **完整支持（保 Refresh）** |
| **EndNote / NoteExpress / Mendeley** | 宏常漏跳或变蓝 | 不支持 | 不支持 | **★ 原生支持（含 Citavi、Word 内置引文）** |
| **无文献软件纯文本手敲引文** | 需逐条手工插书签，极慢 | 不支持 | 不支持 | **★ 原生支持（`[1-3]`、真实上标 `¹⁻³`、中英文作者年份）** |
| **混合稿件（插件引文 + 手改纯文本）** | 极易冲突重复 | 不支持 | 不支持 | **★ 自动协同识别，不漏跳、不重复计数** |
| **定稿导出可点击跳转 PDF** | 满篇蓝字下划线、上标变形 | 不支持（仅软件内跳转） | 支持（限 Zotero 稿件） | **★ 全场景支持导出黑字、无下划线、原小上标可点击 PDF** |
| **科学单位防误连（`m²` / `10⁶` / `R²`）** | 正则宏常把 `m²` 误连到文献 [2] | — | — | **★ 内置学术单位与统计式守卫，精准避开非引文数字** |
| **隐私安全** | — | 全程本机离线解析 | 全程本机离线解析 | **全程本机内存解析，绝不上传论文正文或文献库** |

---

## 二、官方下载（Pro 全能版安装包 + 图文 PDF 教程 + 百度网盘分流）

👉 前往 **[Releases 最新版本页](https://github.com/huangshaoxiang123/RefLink-Pro/releases/latest)** 或直接按下方对应系统与软件下载：

| 对应平台 | GitHub 安装包直链 | 图文 PDF 教程 | 百度网盘极速下载 | 适用环境与说明 |
| :--- | :--- | :--- | :--- | :--- |
| **Mac · WPS 文字（Pro 全能版）** | [RefLink-0.7.2-macOS.zip](https://github.com/huangshaoxiang123/RefLink-Pro/releases/latest/download/RefLink-0.7.2-macOS.zip) (240 KB) | [图文 PDF 教程](https://github.com/huangshaoxiang123/RefLink-Pro/releases/latest/download/RefLink-Pro-Mac-WPS-Guide.pdf) | [网盘下载](https://pan.baidu.com/s/1EH6luRoGJxggh8ZqOmQ7Zg?pwd=8biw)（码 `8biw`） | macOS 12+，支持纯文本 / EndNote / NoteExpress / Zotero，正文/脚注/尾注 + 导出可点击 PDF |
| **Windows · WPS 文字（Pro 全能版）** | [RefLink-WPS-0.7.2-Windows.zip](https://github.com/huangshaoxiang123/RefLink-Pro/releases/latest/download/RefLink-WPS-0.7.2-Windows.zip) (217 KB) | [图文 PDF 教程](https://github.com/huangshaoxiang123/RefLink-Pro/releases/latest/download/RefLink-Pro-Windows-WPS-Guide.pdf) | [网盘下载](https://pan.baidu.com/s/1v0TWLQ5Xmvjk33Ed0afvcQ?pwd=ahve)（码 `ahve`） | Windows 10/11 + WPS 2019+，支持纯文本 / EndNote / NoteExpress / Zotero + 导出可点击 PDF |
| **Mac · Microsoft Word（Pro 全能版）** | [RefLink-Word-0.7.2-macOS.zip](https://github.com/huangshaoxiang123/RefLink-Pro/releases/latest/download/RefLink-Word-0.7.2-macOS.zip) (31 KB) | [图文 PDF 教程](https://github.com/huangshaoxiang123/RefLink-Pro/releases/latest/download/RefLink-Pro-Mac-Word-Guide.pdf) | [网盘下载](https://pan.baidu.com/s/1EqvpnGZGA1mwLwNh8mIrig?pwd=rw8i)（码 `rw8i`） | macOS 12+ + Microsoft 365 / Word 16.x（支持 WordApi 1.5），全文献引擎跳转 + 导出可点击 PDF |
| **Windows · Microsoft Word（Pro 全能版）** | [RefLink-Word-0.7.2-Windows.zip](https://github.com/huangshaoxiang123/RefLink-Pro/releases/latest/download/RefLink-Word-0.7.2-Windows.zip) (10 KB) | [图文 PDF 教程](https://github.com/huangshaoxiang123/RefLink-Pro/releases/latest/download/RefLink-Pro-Windows-Word-Guide.pdf) | [网盘下载](https://pan.baidu.com/s/10pPQJjS3kHJxfzC_a0EQDw?pwd=qgyq)（码 `qgyq`） | Windows 10/11 + Microsoft 365 / 现代 Office 2021（支持 WordApi 1.5），全文献引擎跳转 + 导出可点击 PDF |

> 💡 **国内校园网极速下载（分平台百度网盘）**：若直连 GitHub 较慢，可直接点击对应网盘链接下载同款 `v0.7.2` 安装包与高清图文教程 PDF：
> - **Mac · WPS**：[https://pan.baidu.com/s/1EH6luRoGJxggh8ZqOmQ7Zg?pwd=8biw](https://pan.baidu.com/s/1EH6luRoGJxggh8ZqOmQ7Zg?pwd=8biw) &nbsp;|&nbsp; 提取码：`8biw`
> - **Win · WPS**：[https://pan.baidu.com/s/1v0TWLQ5Xmvjk33Ed0afvcQ?pwd=ahve](https://pan.baidu.com/s/1v0TWLQ5Xmvjk33Ed0afvcQ?pwd=ahve) &nbsp;|&nbsp; 提取码：`ahve`
> - **Mac · Word**：[https://pan.baidu.com/s/1EqvpnGZGA1mwLwNh8mIrig?pwd=rw8i](https://pan.baidu.com/s/1EqvpnGZGA1mwLwNh8mIrig?pwd=rw8i) &nbsp;|&nbsp; 提取码：`rw8i`
> - **Win · Word**：[https://pan.baidu.com/s/10pPQJjS3kHJxfzC_a0EQDw?pwd=qgyq](https://pan.baidu.com/s/10pPQJjS3kHJxfzC_a0EQDw?pwd=qgyq) &nbsp;|&nbsp; 提取码：`qgyq`

---

## 三、3 步快速上手（写作点击跳转 & 导出可点击 PDF）

### 模式 A：日常改稿 · 软件内安全点击跳转
1. 打开论文（无论是纯文本手敲稿、EndNote、NoteExpress 还是 Zotero 稿件），点击顶部 `RefLink 交叉引用` → **启用/更新跳转（WPS）/ 创建/更新（Word）**。
2. 插件自动定位文末参考文献并完成匹配校验；此后在正文点击任意 `[1]`、`[2-3]` 或作者年份引文，光标即刻平滑跳至文末对应文献条目，**全文保持原字体样式不变**。

### 模式 B：论文定稿 · 导出黑字无下划线可点击 PDF
1. **另存定稿副本**：准备提交盲审、答辩或期刊投稿前，先另存一份副本（如 `论文-定稿导出PDF版.docx`）。
2. **生成 PDF 跳转**：在副本中点击 `RefLink 交叉引用` → **生成 PDF 跳转** 并确认（复合角标如 `[2-3]` 会自动分别映射首尾文献，且强制锁定黑色、无下划线、原小上标）。
3. **正确导出 PDF**：
   - **Mac / Windows WPS**：点击 **文件 → 输出为 PDF**；
   - **Mac Microsoft Word**：点击 **文件 → 另存为 → 文件格式选 PDF → 勾选「最适合电子分发和辅助功能」**（切勿使用系统“打印 → 存储为 PDF”）；
   - **Windows Microsoft Word**：点击 **文件 → 导出 → 创建 PDF/XPS**（切勿使用“Microsoft Print to PDF”虚拟打印）。

---

## 四、先免费检测，满意再激活（Pro 全能版 `RLKX-` 离线永久授权）

1. **免费只读检测**：安装后无需激活即可对自己的论文运行 **「检查文档（WPS）/ 检测（Word）」**，直接查看识别到的参考文献条数、引文来源类型（如 `纯文本引文`、`EndNote 引文域`、`Zotero 引文域`）与可跳转片段总数。
2. **一机一码永久有效**：点击功能按钮查看**本机机器码**（`RLD-XXXX-XXXX-XXXX`），联系作者获取 **Pro 全能版永久激活码（`RLKX-` 开头）** 或 `.txt` 激活文件（Mac WPS 支持一键导入 `.txt` 自动激活）。
3. **老用户补差价丝滑升级**：已购标准版（`RLK1-`）或 Plus 版（`RLKP-`）的用户，安装 `v0.7.2` 后凭同台设备机器码联系作者补差价获取 `RLKX-` 激活码，粘贴或导入即可原地升级至 Pro 全能版！

---

## 五、获取 Pro 永久激活码与技术支持

- **作者微信**：`hsx15299177861`（添加请备注 **RefLink Pro**）
- **问题反馈**：欢迎在本仓库 [Issues](https://github.com/huangshaoxiang123/RefLink-Pro/issues) 提交兼容性反馈或建议。

---

## 版权与许可声明

本软件（RefLink Pro）为**闭源商业授权软件（All Rights Reserved）**，详见 [LICENSE.txt](./LICENSE.txt)。本 GitHub 仓库仅用于官方安装包分发、使用文档展示与用户问题反馈，**未经作者书面许可，禁止对安装包进行逆向工程、反混淆、修改、去授权、二次打包或商业转售**。

> “Zotero”、“EndNote”、“NoteExpress”、“Mendeley”、“Microsoft Word” 与 “WPS” 均为各自权利人的商标或注册商标。RefLink 为独立第三方科研效率工具，与上述官方无隶属或背书关系。
