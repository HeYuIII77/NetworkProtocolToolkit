# Architecture

## 架构风格
单文件 WinForms 应用。Form1.Designer.cs 只含设计器基本属性，所有 UI 创建和业务逻辑在 Form1.cs。

## 代码组织

### Form1.Designer.cs (~50行)
- InitializeComponent(): 窗体基本属性（Size、Text、FormBorderStyle 等）
- 无控件创建代码，设计器可正常加载

### Form1.cs (2500+ 行)
- **构造函数**: InitializeComponent → InitializeUI（Designtime 跳过）→ 日志初始化
- **InitializeUI()**: 构建完整 UI — 左侧 ListBox 导航 + 右侧隐藏 TabControl
- **Create*Tab()**: 13 个页面创建方法（CreateNetworkDiagTab / CreateConnectivityTab / ...）
- **LabeledTextBox()**: 带标签的文本框辅助方法
- **日志系统**: EnsureLogDirectory / WriteToDailyLog / AppendDetailedLog / AppendLog
- **网络诊断方法**: DoPing / DoDnsLookup / DoPortScan / DoSslCertCheck
- **Web 协议方法**: DoHttpRequest（通用）/ DoHttpGet / DoHttpPost / DoRestPost / DoWebSocketEcho / DoWebServiceTest
- **文件传输方法**: DoFtpList / DoFtpUpload / DoFtpDownload / DoFtpDelete / DoFtpMkdir / DoSftpList_Strong
- **SSH 方法**: DoSshCommand
- **邮件方法**: DoSendSmtp / DoPop3MailKit / DoImapMailKit
- **连通性方法**: TestIpPort / TestDbConnection
- **工业协议方法**: TestS7ReadWrite / TestModbusTcp / TestModbusWrite / TestRawTcp / DoOpcUaRead / DoOpcUaWrite / DoOpcDaReadWrite
- **连接字符串构建**: BuildSqlServerConnectionString / BuildOracleConnectionString / BuildMySqlConnectionString / BuildNpgsqlConnectionString
- **辅助方法**: AddDbResultRow / MaskConnectionString / FirstChars / OpenLogFolder / ExportTodayLog

## UI 布局
```
┌──────────────────────────────────────┐
│ 窗体 (FixedSingle, 1200x820)         │
├──────────────┬───────────────────────┤
│ Panel(Left)  │ Panel(Fill)           │
│ Width=600    │                       │
│ ┌──────────┐ │ ┌───────────────────┐ │
│ │Header 40px│ │ │TabControl(Fill)  │ │
│ ├──────────┤ │ │(FlatButtons,      │ │
│ │ListBox   │ │ │ ItemSize=(0,1))   │ │
│ │(Fill)    │ │ │                   │ │
│ │13 个导航项│ │ │ 各 Tab 页面内容    │ │
│ └──────────┘ │ └───────────────────┘ │
└──────────────┴───────────────────────┘
```

## 线程模型
- 所有网络操作使用 `async/await` 异步执行
- 端口扫描使用 SemaphoreSlim(80) 并发控制 + CancellationTokenSource
- UI 更新通过 BeginInvoke 或 WinForms 自动同步
- 日志文件写入使用 `lock (_logLock)` 同步

## 依赖包
| 包名 | 版本 | 用途 |
|------|------|------|
| MailKit | 4.14.1 | POP3/IMAP |
| Microsoft.Data.SqlClient | 6.1.2 | SQL Server |
| MySql.Data | 9.4.0 | MySQL |
| Npgsql | 8.0.3 | PostgreSQL |
| OPCFoundation.NetStandard.Opc.Ua | 1.5.377.21 | OPC UA |
| Renci.SshNet.Async | 1.4.0 | SFTP / SSH |
| S7netplus | 0.20.0 | S7 协议 |

## 发布配置
- `PublishSingleFile`: true
- `SelfContained`: true
- `RuntimeIdentifier`: win-x64
- `FormBorderStyle`: FixedSingle（不可调整大小）
- `MaximizeBox`: false
