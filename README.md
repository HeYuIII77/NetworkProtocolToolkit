# 网络协议工具箱

## 概述
该应用程序是一个基于 **Windows Forms (.NET 8.0)** 的综合网络协议测试和调试工具，旨在帮助开发人员、系统集成工程师和网络管理员快速测试各种网络协议的连通性和功能性。  
工具支持从基础的 **TCP/IP 连通性测试** 到复杂的 **工业协议通信**，提供了统一的界面来执行和监控各种网络操作。  
程序需要 **管理员权限** 才能修改某些系统设置和访问受限资源。

---

## 功能

### 网络诊断
- **Ping**：ICMP 连通性测试，支持自定义次数
- **DNS 解析**：域名解析查询
- **端口扫描**：并发端口扫描，支持实时进度反馈和停止按钮
- **SSL/TLS 证书检查**：查看目标站点证书有效期和详情

### 连通性检测
- **IP+端口可达性测试**（TCP 连接测试）
- **多种数据库连接测试**（SQL Server、Oracle、MySQL、PostgreSQL）
- 实时结果显示和耗时统计

### Web 协议测试
- **HTTP/HTTPS**：支持 GET/POST/PUT/DELETE/PATCH，自定义 Headers，Basic/Bearer 认证
- **REST API**：POST JSON 请求测试
- **WebSocket**：连接测试和消息回显
- **WebService (SOAP)**：SOAP 协议调用，支持自定义 SOAPAction

### 文件传输协议
- **FTP**：目录列表、上传、下载、删除、创建目录（匿名/认证）
- **SFTP**：目录列表，支持密码和私钥认证（Renci.SshNet）
- **SSH**：远程命令执行

### 邮件协议
- **SMTP**：邮件发送功能，支持 SSL/TLS
- **POP3/IMAP**：邮件收取功能（MailKit），支持 SSL/TLS

### 工业协议测试
- **S7 协议**：西门子 PLC 通信，支持读写操作（S7netplus）
- **Modbus TCP**：寄存器读取 + 写操作（FC05/FC06/FC16）
- **OPC UA**：节点读写（OPCFoundation 官方库）
- **OPC DA**：通过 COM 互操作支持
- **Raw TCP**：原始 TCP 通信测试

### 辅助功能
- 分级日志系统（INFO/WARN/ERROR），按日期分割
- 日志导出和文件夹管理
- 连接字符串敏感信息隐藏
- 动态检测可用的数据库驱动，提供 NuGet 安装建议

---

## 界面布局

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

- **左侧 ListBox 导航**：13 个功能模块，点击切换右侧内容
- **统一控件风格**：LabeledTextBox 辅助方法创建一致的输入控件
- **响应显示区域**：每个协议页面都有专门的响应显示文本框

---

## 文件结构

```
network-protocol-toolkit/
├── NetworkProtocolToolkit/
│   ├── Form1.cs              # 主窗体业务逻辑（2500+ 行）
│   ├── Form1.Designer.cs     # UI 设计器代码（仅基本属性）
│   ├── Program.cs            # 入口点
│   ├── NetworkProtocolToolkit.csproj
│   └── app.manifest          # 管理员权限声明
├── NetworkProtocolToolkit.sln
├── README.md
└── LICENSE
```

### NuGet 包

| 包名 | 版本 | 用途 |
|------|------|------|
| MailKit | 4.14.1 | POP3/IMAP 邮件收取 |
| Microsoft.Data.SqlClient | 6.1.2 | SQL Server 连接 |
| MySql.Data | 9.4.0 | MySQL 连接 |
| Npgsql | 8.0.3 | PostgreSQL 连接 |
| OPCFoundation.NetStandard.Opc.Ua | 1.5.377.21 | OPC UA 通信 |
| Renci.SshNet.Async | 1.4.0 | SFTP / SSH |
| S7netplus | 0.20.0 | 西门子 S7 协议 |

---

## 运行环境

- **.NET 8.0**（Windows Forms）
- **Windows 10/11**（win-x64）
- **需要管理员权限** 以访问某些系统资源和修改环境变量

---

## 发布与使用

### 自包含发布（推荐）

项目支持自包含单文件发布，无需安装 .NET 运行时：

```bash
dotnet publish NetworkProtocolToolkit/NetworkProtocolToolkit.csproj \
  -c Release -r win-x64 \
  --self-contained true \
  -p:PublishSingleFile=true \
  -p:EnableCompressionInSingleFile=true \
  -o ./publish
```

发布产物：
- `NetworkProtocolToolkit.exe`（约 75MB，包含所有依赖）
- 解压 zip 后双击直接运行，不需要额外下载

### 基本操作流程

1. **启动应用程序** — 程序自动创建日志目录并记录启动信息
2. **选择测试协议** — 通过左侧导航栏选择要测试的协议类型
3. **执行测试** — 填写连接参数，点击测试按钮，查看结果
4. **查看日志** — 切换到"日志"页面查看详细操作记录

---

## 注意事项

### 权限要求
- **管理员权限**：某些操作需要管理员权限
- **防火墙配置**：测试外部服务时可能需要配置防火墙规则

### 安全考虑
- 连接字符串中的密码信息会在日志中自动隐藏
- HTTPS 测试时可能需要处理自签名证书
- 工业协议测试通常在隔离的网络环境中进行

### 性能提示
- 所有网络操作使用 async/await 异步执行，不阻塞 UI
- 端口扫描支持并发控制和手动停止
- 所有网络连接和资源都会正确释放

---

## 故障排除

| 问题 | 解决方案 |
|------|---------|
| 未找到数据库驱动 | 通过 NuGet 安装对应的数据库连接包 |
| 连接超时 | 检查目标服务是否运行，验证网络连通性和防火墙设置 |
| 证书验证失败 | 测试环境可暂时忽略证书验证，或安装 CA 证书 |
| 权限不足 | 以管理员身份运行程序 |

---

该工具箱持续更新，欢迎反馈使用体验和改进建议！
