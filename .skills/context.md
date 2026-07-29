# Current Development Context

**Last updated**: 2026-07-29

## Currently Working On
- 无活跃开发任务

## Current Module
- UI 布局已稳定为 ListBox 左侧导航方案

## Completed This Session
- 修复输入框标题过长被遮挡：缩短 5 个标题 + 2 个改为多行显示
- 修复 Label 高度不足导致文字底部被裁：Panel 30→34、Label 20→26、Top 4→2
- 修复 HTTP 认证下拉框被遮挡：authTypePanel 高度对齐
- 统一所有硬编码 Height=30 的 Panel（SMTP SSL、Modbus FC、POP3 SSL、IMAP SSL）
- LabeledTextBox 方法新增 labelMultiline 参数支持多行标题
- 发布自包含单文件 exe（75MB）+ zip 分发包（74MB），包含所有 .NET 运行时和 NuGet 依赖

## Not Yet Complete
- None

## Blockers
- None

## Next Steps
- 暂无具体计划
- 可选：拆分 Form1.cs 单文件、添加单元测试
