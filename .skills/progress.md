# Development Progress

## 2025-10-28

Completed:
- 初始项目创建
- 添加 .gitattributes / .gitignore / LICENSE
- 添加项目文件

## 2025-10-29

Completed:
- 补充代码注释
- 添加 README.md 文档
- 删除多余文件

## 2026-07-28

Completed:
- 初始化 `.skills/` 项目状态管理
- 全面分析项目结构和代码状态
- **新增"网络诊断"Tab**：Ping (ICMP)、DNS 解析、端口扫描（并发）、SSL/TLS 证书检查
- **增强 HTTP/HTTPS Tab**：PUT/DELETE/PATCH 方法、自定义 Headers、Basic/Bearer 认证、响应头显示
- **增强 FTP/SFTP Tab**：FTP 上传/下载/删除/创建目录、SSH 远程命令执行（Renci.SshNet）
- **增强 Modbus TCP**：FC05 写线圈、FC06 写单寄存器、FC16 写多寄存器
- **新增 PostgreSQL 支持**：添加 Npgsql 8.0.3 包、连接字符串构建、连通性检测按钮
- **修复端口扫描卡 UI**：实时进度反馈、停止按钮、CancellationTokenSource 取消机制
- **修复 OPC UA NullReferenceException**：Designer 中 OPC UA 按钮事件漏赋值 _opcUaRespBox
- **UI 布局重构**：从 TabControl 改为 ListBox 左侧导航 + 隐藏 TabControl 内容容器
- **设计器兼容**：InitializeComponent 精简为基本属性，InitializeUI() 移到 Form1.cs，LicenseManager.Designtime 检查
- **导航面板方案**：从 SplitContainer 改为双 Panel（Dock=Left 600px + Dock=Fill）

Added:
- `NetworkProtocolToolkit.csproj`: Npgsql 8.0.3 包引用
- `Form1.cs`: 约 800 行新代码（25+ 个新方法 + InitializeUI + Create*Tab 方法）
- `Form1.Designer.cs`: 精简为 50 行基本属性

Modified:
- `Form1.cs`: 构造函数增加 InitializeUI() 调用 + Designtime 判断
- `Form1.Designer.cs`: 从 1000+ 行精简为 50 行（所有 UI 创建移到 Form1.cs）

Decisions:
- Tab 布局最终选择 ListBox 左侧导航（放弃 OwnerDraw 垂直 Tab、旋转画布等方案）
- 设计器兼容方案：InitializeComponent 只放基本属性，复杂逻辑放 InitializeUI()
