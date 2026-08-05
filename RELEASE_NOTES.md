# 群众投票 v0.3.6

本版本提供 Windows 10/11 64 位无外部依赖便携版。完整解压 ZIP 后即可运行，不需要安装 Python、Microsoft Office、WPS、LibreOffice或其他运行库。

## 主要更新

- 新增 WPS Writer 自动化支持：没有 Microsoft Word 但安装了 WPS 时，优先由 WPS 生成真实 PDF 打印预览。
- 转换顺序固定为 Microsoft Word → WPS Writer → 随便携版内置的 LibreOffice Writer 引擎。
- Word/WPS 自动化改为独立后台进程，并加入 120 秒超时；异常、无效注册或导出失败时自动清理临时 PDF，超时时结束完整辅助进程树并回退，不阻塞主界面。
- 未安装 Word/WPS 时继续由内置引擎生成真实 PDF 打印预览，纯净 Windows 电脑仍可直接运行。
- `.doc` 模板转换也改为使用内置引擎。
- 每次转换使用独立临时用户配置，避免并发预览、快速微调时发生文件锁和进程冲突。
- 增加转换超时、隐藏后台窗口和明确的运行时缺失提示。
- 保留文字和打勾拖拽、方向键微调、模板判断区/标记区自动保存恢复、批量 DOCX 导出等现有功能。
- 清理旧缓存、旧构建产物和仓库示例数据；正式便携包不包含任何示例或用户数据。

## 使用方法

下载 `QunzhongVote-v0.3.6-Windows-Portable.zip`，完整解压后双击 `QunzhongVote.exe`（启动后为中文界面）。请不要只复制 EXE，旁边的 `runtime` 文件夹是程序自带的文档引擎。

模板配置仍保存在本机 `%LOCALAPPDATA%\QunzhongVote\template_profiles.json`，升级不会删除已保存的判断区和标记区。
