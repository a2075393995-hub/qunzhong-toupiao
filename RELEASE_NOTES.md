# 群众投票 v0.3.4

本版本将正式更新源迁移到 Gitee，并提供带真实截图、脱敏测试数据和 PDF 使用说明的完整交付包。

## 更新内容

- 应用内“检查更新”从 GitHub 迁移到公开的 Gitee Release API，下载页和源码仓库入口均指向 Gitee。
- Gitee API 不返回 `html_url` 时，客户端会根据 Release 标签生成正确的 Gitee 下载页地址。
- Gitee OpenAPI 不可用或受限时，会从公开 Releases 页面读取最新标签作为备用。
- 新增 Gitee API、Release 页面回退和下载页地址单元测试。
- 新增完全虚构的测试数据，不包含用户姓名、真实电话或原始投票数据。
- 使用说明升级为带截图的 PDF，覆盖六步流程、模板标注、真实打印预览、位置微调、导出模式和常见问题。
- 延续 v0.3.3 的真实 Word→PDF 预览、文字/打勾拖拽、方向键微调、线程安全刷新和模板标注恢复功能。

## 下载与使用

下载 `qunzhong-toupiao.exe` 后直接运行。模板配置保存在本机 `%LOCALAPPDATA%\QunzhongVote\template_profiles.json`，模板、投票数据、配置和导出结果都不会上传到网络。Gitee Release 同时提供 PDF 使用说明和脱敏测试数据。
