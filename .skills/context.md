# Current Development Context

**Last updated**: 2026-07-28

## Currently Working On
- 无活跃开发任务

## Current Module
- UI 布局已稳定为 ListBox 左侧导航方案

## Completed This Session
- 新增"网络诊断"Tab：Ping、DNS 解析、端口扫描、SSL 证书检查
- 增强 HTTP/HTTPS Tab：PUT/DELETE/PATCH + 自定义 Headers + Basic/Bearer 认证
- 增强 FTP/SFTP Tab：FTP 上传/下载/删除/创建目录 + SSH 远程命令执行
- 增强 Modbus TCP：FC05/FC06/FC16 写操作
- 新增 PostgreSQL 数据库连接支持（Npgsql 8.0.3）
- 修复端口扫描卡 UI：实时进度反馈 + 停止按钮 + CancellationTokenSource
- 修复 OPC UA NullReferenceException：_opcUaRespBox 未赋值
- UI 布局多次迭代：顶部水平 Tab → 左侧垂直 Tab(OwnerDraw) → 旋转画布 → DirectionVertical → 顶部多行 → 最终 ListBox 左侧导航
- 修复设计器兼容性：InitializeComponent 精简为基本属性，InitializeUI 移到 Form1.cs，Designtime 跳过
- 导航面板宽度从 SplitContainer 改为双 Panel 方案（Dock=Left + Dock=Fill）

## Not Yet Complete
- None

## Blockers
- None

## Next Steps
- 暂无具体计划
- 可选：拆分 Form1.cs 单文件、添加单元测试
