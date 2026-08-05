# 群众选票格式化打印工具

面向 Windows 的群众选票格式化打印工具。读取 Word 选票模板和 CSV/XLSX 投票数据，通过可视化标注写入姓名、房号、电话和投票结果，批量生成 DOCX，并输出投票结果汇总表。

当前版本：v0.3.5。正式发布物为无外部依赖便携版，目标电脑不需要安装 Python、Microsoft Office、LibreOffice或其他运行库。应用内更新源为 Gitee，GitHub 作为同步镜像和自动构建渠道。

## 功能

- 导入 `.docx`/`.doc` 模板和 CSV/XLSX 数据。
- 可视化设置判断区、标记区和用户字段位置。
- 真实 PDF 打印预览，显示纸张尺寸、方向和分页。
- 楼栋、房号、完整地址、姓名、电话和打勾均可拖拽；方向键移动 1pt、Shift 5pt、Ctrl 0.1pt。
- 精确预览刷新完成前禁止确认导出，避免旧画面被误确认。
- 模板判断区、标记区、字段位置和样式按模板内容自动保存并在重启后恢复。
- 支持单文件和多文件 DOCX 导出以及投票汇总表。

## 无依赖便携版

运行 `tools/build_portable.ps1` 会从 Document Foundation 官方下载 LibreOffice MSI，通过管理提取方式准备便携运行时，不会安装到构建电脑。发布包结构如下：

```text
QunzhongVote-v0.3.5-Portable/
├─ QunzhongVote.exe（启动后为中文界面）
├─ PORTABLE_README.txt
└─ runtime/libreoffice/...
```

目标电脑完整解压后直接运行 EXE。电脑存在 Microsoft Word 时程序优先使用 Word；没有 Word 时自动调用便携包内置引擎。每次转换使用独立临时配置目录，不读取或修改目标电脑上的 LibreOffice 用户配置。

## 本地开发

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt -r requirements-dev.txt
python -m unittest discover -s tests -v
python app.py
```

构建完整便携版：

```powershell
.\tools\build_portable.ps1
```

正式便携 ZIP 输出到 `dist/QunzhongVote-v0.3.5-Windows-Portable.zip`。

## 代码结构

- `app.py`：桌面界面和工作流。
- `vote_core.py`：模板解析、标记写入和批量导出。
- `office_runtime.py`：内置文档引擎定位、隔离配置和转换执行。
- `print_preview.py`：DOCX→PDF、纸张识别和 PDF 页面渲染。
- `template_profiles.py`：模板配置持久化。
- `update_service.py`：Gitee 更新检查。
- `tools/prepare_libreoffice.ps1`：下载并准备官方便携运行时。
- `tools/build_portable.ps1`：测试、构建和压缩 Windows 便携版。

## 数据安全

仓库和正式发布包不包含示例投票数据、真实模板或用户数据。模板、投票数据、预览和导出结果默认仅保存在本机，联网只用于用户主动检查 Gitee 更新。
