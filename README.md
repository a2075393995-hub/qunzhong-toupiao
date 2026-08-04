# 群众投票

一个面向 Windows 的群众选票格式化打印工具。程序读取 Word 选票模板和 CSV/XLSX 投票数据，通过可视化标注将姓名、房号、电话和投票结果写入模板，批量生成 DOCX，并输出投票结果汇总表。

## 主要功能

- 导入 Word 模板和 CSV/XLSX 数据。
- 可视化调试字段位置和投票打勾位置。
- 导出前预览，确认后再批量生成文档。
- 支持一人一份 DOCX 或合并为单个 DOCX。
- 支持纯净打印模式、异常数据检查和 XLSX 汇总。
- 内置“开发更新”按钮，从 GitHub Releases 检查新版本。

## 本地运行

需要 Python 3.10 或更高版本。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

## 构建 Windows EXE

```powershell
python -m pip install -r requirements.txt -r requirements-dev.txt
python -m PyInstaller --noconfirm --clean VoteDocxApp.spec
```

构建结果位于 `dist/QunzhongVote.exe`。也可以直接运行 `build_exe.bat`，脚本会生成带版本号的中文文件名。

## 发布和应用内更新

版本号定义在 `update_service.py` 的 `APP_VERSION`。

1. 修改 `APP_VERSION`，例如从 `0.1.0` 改为 `0.2.0`。
2. 提交代码并推送标签 `v0.2.0`。
3. GitHub Actions 自动构建 Windows EXE，并创建对应 Release。
4. 已安装的程序点击“开发更新”后，会匿名访问公开的 GitHub Releases API，比对版本并打开新版下载页。

客户端不需要 GitHub Token，也不会保存或上传用户的选票模板、投票数据和导出文件。

## 项目结构

- `app.py`：Tkinter 桌面界面和操作流程。
- `vote_core.py`：数据解析、模板标注和 DOCX/XLSX 导出核心。
- `update_service.py`：GitHub Release 检查与版本比较。
- `qa/`：模板和界面回归辅助脚本。
- `tests/`：可独立运行的单元测试。
- `.github/workflows/release.yml`：按版本标签自动构建和发布。

## 数据与隐私

模板、投票数据、预览和导出结果默认只保存在本机。仓库忽略 `output/`、构建目录、缓存以及常见本地环境文件；提交前仍应自行确认测试数据不含真实个人信息。
