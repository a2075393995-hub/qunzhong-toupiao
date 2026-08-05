# 群众投票

一个面向 Windows 的群众选票格式化打印工具。程序读取 Word 选票模板和 CSV/XLSX 投票数据，通过可视化标注将姓名、房号、电话和投票结果写入模板，批量生成 DOCX，并输出投票结果汇总表。

当前版本：v0.3.4。判断区、标记区、字段位置和样式会按模板内容自动保存在本机；重新选择同一模板或重启程序后会自动恢复。应用内更新源为 Gitee，GitHub 作为同步镜像和自动构建渠道。

## 主要功能

- 导入 Word 模板和 CSV/XLSX 数据。
- 可视化设置字段位置和投票打勾位置。
- 使用 Microsoft Word 导出 PDF 的真实打印预览，自动显示纸张尺寸、方向和分页。
- 所有最终位置调整统一放在真实打印预览中：楼栋、房号、完整地址、姓名、电话和当前票面的打勾都可直接拖动，也可用方向键微调（普通 1pt、Shift 5pt、Ctrl 0.1pt），停手后自动校准为精确的 Word→PDF 效果。
- 导出前预览只显示 Word 直接导出的真实 PDF 页面；精确刷新完成前禁止确认导出，避免预览与打印结果不一致。
- 自动把 `1-101`、`1栋101室` 等地址拆成楼栋与房号，并保留模板原有的“栋/幢/室”和下划线格式。
- 支持一人一份 DOCX 或合并为单个 DOCX。
- 支持纯净打印模式、异常数据检查和 XLSX 汇总。
- 内置“检查更新”按钮，从 Gitee Releases 检查正式版本并打开 Gitee 下载页。
- 仓库内提供完全虚构的 `samples/示例投票数据.csv`，可用于功能测试和异常票验证。

## 本地运行

需要 Python 3.10 或更高版本。

真实打印预览优先使用本机 Microsoft Word；未安装 Word 时可安装 LibreOffice 作为备用转换器。

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
3. GitHub Actions 自动构建 Windows EXE，并创建 GitHub 镜像 Release；同一 EXE 同步上传到 Gitee Release。
4. 已安装的程序点击“检查更新”后，会匿名访问公开的 Gitee Releases，比对版本并打开新版下载页。

客户端不需要 Gitee Token，也不会保存或上传用户的选票模板、投票数据和导出文件。

- Gitee 主更新仓库：`https://gitee.com/zhang-jiaxin654/qunzhong-toupiao`
- GitHub 同步镜像：`https://github.com/a2075393995-hub/qunzhong-toupiao`

## 项目结构

- `app.py`：Tkinter 桌面界面和操作流程。
- `vote_core.py`：数据解析、模板标注和 DOCX/XLSX 导出核心。
- `print_preview.py`：Word→PDF 固定版式导出、纸张识别和 PDF 页面渲染。
- `update_service.py`：Gitee Release 检查与版本比较。
- `qa/`：模板和界面回归辅助脚本。
- `tests/`：可独立运行的单元测试。
- `.github/workflows/release.yml`：按版本标签自动构建和发布。

## 数据与隐私

模板、投票数据、预览和导出结果默认只保存在本机。仓库忽略 `output/`、构建目录、缓存以及常见本地环境文件；提交前仍应自行确认测试数据不含真实个人信息。
