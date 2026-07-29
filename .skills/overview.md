# Project Overview

## 项目定位
基于 Windows Forms 的综合网络协议测试和调试工具，帮助开发人员、系统集成工程师和网络管理员快速测试各种网络协议的连通性和功能性。

## 目标用户
- 开发人员
- 系统集成工程师
- 网络管理员

## 技术栈

| 层级 | 技术 |
|------|------|
| 框架 | .NET 8.0 (Windows Forms) |
| 平台 | win-x64 |
| 语言 | C# |
| UI | Windows Forms + TabControl |

## 核心依赖

| 包名 | 版本 | 用途 |
|------|------|------|
| MailKit | 4.14.1 | POP3/IMAP 邮件收取 |
| Microsoft.Data.SqlClient | 6.1.2 | SQL Server 连接 |
| MySql.Data | 9.4.0 | MySQL 连接 |
| OPCFoundation.NetStandard.Opc.Ua | 1.5.377.21 | OPC UA 通信 |
| Renci.SshNet.Async | 1.4.0 | SFTP 功能 |
| S7netplus | 0.20.0 | 西门子 S7 协议 |

## 目录结构

```
network-protocol-toolkit/
├── NetworkProtocolToolkit/
│   ├── Form1.cs              # 主窗体业务逻辑 (1519行)
│   ├── Form1.Designer.cs     # UI 设计器代码 (738行)
│   ├── Program.cs            # 入口点
│   ├── NetworkProtocolToolkit.csproj
│   └── app.manifest          # 管理员权限声明
├── NetworkProtocolToolkit.sln
├── README.md
└── LICENSE
```

## 功能模块

### 连通性检测
- IP+端口可达性测试 (TCP)
- 数据库连接测试 (SQL Server / MySQL)

### Web 协议
- HTTP GET/POST
- REST API POST JSON
- WebSocket 回显
- WebService SOAP

### 文件传输
- FTP 目录列表（匿名/认证）
- SFTP 目录列表（密码/私钥）

### 邮件协议
- SMTP 发送（SSL/TLS）
- POP3 收取（MailKit）
- IMAP 收取（MailKit）

### 工业协议
- S7 读写（西门子 PLC）
- Modbus TCP 寄存器读取
- OPC UA 节点读写
- OPC DA 读写（COM 互操作）
- Raw TCP 原始通信

### 辅助功能
- 分级日志系统（INFO/WARN/ERROR）
- 按日期分割日志文件
- 日志导出与文件夹管理
- 敏感信息隐藏
