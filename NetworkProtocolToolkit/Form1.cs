using System.Net;
using System.Net.Mail;
using System.Net.Sockets;
using System.Net.WebSockets;
using System.Net.Security;
using System.Net.NetworkInformation;
using System.Reflection;
using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using Opc.Ua;
using Opc.Ua.Client;
using Opc.Ua.Configuration;
using S7.Net;

// SFTP (Renci.SshNet) and MailKit (POP3/IMAP)
using Renci.SshNet;
using Renci.SshNet.Sftp;
using MailKit;
using MailKit.Net.Pop3;
using MailKit.Net.Imap;
using MailKit.Security;
using MimeKit;
using System.ComponentModel;
using System.Diagnostics;

namespace NetworkProtocolToolkit
{
    /// <summary>
    /// 主窗体 - 协议调试工具箱窗体实现（部分）。
    /// 该文件包含窗体的业务逻辑实现：HTTP/REST/WebSocket/FTP/SFTP/SMTP/POP3/IMAP、设备协议（S7/Modbus/Raw TCP/OPC）等功能。
    /// Designer 相关的控件创建由 Form1.Designer.cs 管理；此处只使用那些由 Designer 注入的控件字段（如 _logBox、_httpResponseBox 等）。
    /// </summary>
    public partial class Form1 : Form
    {
        // 日志显示控件（在 Designer 中创建）
        private TextBox _logBox;

        // 共享的 HttpClient 实例，用于所有 HTTP/REST/WebService 请求（推荐重用以节省连接开销）
        private HttpClient _httpClient = new();

        // 各功能页的响应显示控件（由 Designer 初始化）
        private TextBox _httpResponseBox;
        private TextBox _httpHeaderBox;
        private TextBox _restResponseBox;
        private DataGridView _dbResultGrid;
        private TextBox _deviceProtoResponseBox;
        private TextBox _opcUaEndpointBox;
        private TextBox _opcUaNodeBox;
        private TextBox _opcUaRespBox;
        private TextBox _opcDaProgIdBox;    
        private TextBox _opcDaHostBox;
        private TextBox _opcDaRespBox;
        private TextBox _wsResponseBox;
        private TextBox _ftpResponseBox;
        private TextBox _sshResponseBox;

        // 日志内存与目录（线程安全访问）
        private readonly List<string> _logLines = new();
        private readonly string _logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? ".", "Log");
        private readonly object _logLock = new();

        /// <summary>
        /// 构造函数：初始化组件并准备日志目录。
        /// 注意：不要在 Designer 文件中放置业务逻辑，事件处理放在 Form1.cs（当前文件）。
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            // 设计模式下不执行自定义 UI 初始化，否则设计器会报错
            if (LicenseManager.UsageMode != LicenseUsageMode.Designtime)
            {
                InitializeUI();
                EnsureLogDirectory();
                AppendLog("应用启动", "INFO");
            }
        }

        #region UI 初始化

        /// <summary>
        /// 构建完整 UI：左侧 ListBox 导航 + 右侧隐藏标签的 TabControl。
        /// 此方法从 InitializeComponent() 分离出来，使设计器能正常加载。
        /// </summary>
        private void InitializeUI()
        {
            const int navWidth = 300;

            // === 左侧导航面板 ===
            var leftPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = navWidth,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 245)
            };

            var navHeader = new Label
            {
                Text = "协议工具箱",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new System.Drawing.Font("Microsoft YaHei", 12f, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                BackColor = System.Drawing.Color.FromArgb(50, 60, 80),
                ForeColor = System.Drawing.Color.White
            };

            var navList = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Microsoft YaHei", 11f),
                IntegralHeight = false,
                ItemHeight = 30,
                BorderStyle = BorderStyle.None
            };

            var navItems = new[]
            {
                "📡  网络诊断",
                "🔗  连通性检测",
                "📤  REST 接口测试",
                "🌐  HTTP / HTTPS",
                "🔌  WebSocket",
                "📄  WebService (SOAP)",
                "📁  FTP / SFTP / SSH",
                "✉️  SMTP 发送邮件",
                "🏭  设备协议测试",
                "📬  收取邮件 (MailKit)",
                "📋  日志",
                "❓  帮助",
                "📝  更新日志"
            };
            foreach (var item in navItems) navList.Items.Add(item);

            leftPanel.Controls.Add(navList);
            leftPanel.Controls.Add(navHeader);

            // === 右侧内容面板 ===
            var rightPanel = new Panel
            {
                Dock = DockStyle.Fill
            };

            var tab = new TabControl
            {
                Dock = DockStyle.Fill,
                Appearance = TabAppearance.FlatButtons,
                ItemSize = new System.Drawing.Size(0, 1),
                SizeMode = TabSizeMode.Fixed
            };

            tab.TabPages.Add(CreateTab("网络诊断", CreateNetworkDiagTab));
            tab.TabPages.Add(CreateTab("连通性检测", CreateConnectivityTab));
            tab.TabPages.Add(CreateTab("REST 接口", CreateRestTab));
            tab.TabPages.Add(CreateTab("HTTP/HTTPS", CreateHttpTab));
            tab.TabPages.Add(CreateTab("WebSocket", CreateWebSocketTab));
            tab.TabPages.Add(CreateTab("WebService", CreateWebServiceTab));
            tab.TabPages.Add(CreateTab("FTP/SFTP/SSH", CreateFtpSftpTab));
            tab.TabPages.Add(CreateTab("SMTP", CreateSmtpTab));
            tab.TabPages.Add(CreateTab("设备协议", CreateDeviceProtocolsTab));
            tab.TabPages.Add(CreateTab("收取邮件", CreateMailReceiveTab));
            tab.TabPages.Add(CreateTab("日志", CreateLogTab));
            tab.TabPages.Add(CreateTab("帮助", CreateHelpTab));
            tab.TabPages.Add(CreateTab("更新日志", CreateUpdateLogTab));

            navList.SelectedIndexChanged += (_, __) =>
            {
                if (navList.SelectedIndex >= 0 && navList.SelectedIndex < tab.TabCount)
                    tab.SelectedIndex = navList.SelectedIndex;
            };
            navList.SelectedIndex = 0;

            rightPanel.Controls.Add(tab);

            // 先加 Fill 的，再加 Left 的，保证左侧面板不被挤压
            Controls.Add(rightPanel);
            Controls.Add(leftPanel);
        }

        /// <summary>
        /// 辅助方法：将工厂方法返回的控件包装为 TabPage。
        /// </summary>
        private TabPage CreateTab(string title, Func<Control> contentGenerator)
        {
            var content = contentGenerator();

            if (content is TabPage existingPage)
            {
                if (string.IsNullOrWhiteSpace(existingPage.Text)) existingPage.Text = title;
                return existingPage;
            }

            var page = new TabPage(title);
            content.Dock = DockStyle.Fill;
            page.Controls.Add(content);
            return page;
        }

        /// <summary>
        /// 辅助方法：创建带标签的文本框。
        /// </summary>
        private (Panel Panel, TextBox TextBox) LabeledTextBox(string label, string defaultText, bool multiline = false, int height = 34, bool labelMultiline = false)
        {
            int panelHeight = height;
            if (labelMultiline && height < 48) panelHeight = 48; // 多行 Label 至少 48px
            var panel = new Panel { Width = 940, Height = panelHeight };
            var lbl = new Label
            {
                Text = label,
                Left = 0,
                Top = 2,
                Width = 200,
                AutoSize = false,
                Height = labelMultiline ? panelHeight - 2 : 26
            };
            var tbTop = labelMultiline ? (panelHeight - Math.Max(16, height - 6)) / 2 : 0;
            var tb = new TextBox
            {
                Left = 210,
                Top = tbTop,
                Width = 700,
                Height = Math.Max(16, height - 6),
                Text = defaultText,
                Multiline = multiline,
                ScrollBars = multiline ? ScrollBars.Both : ScrollBars.None,
                Font = new System.Drawing.Font("Consolas", 10)
            };
            panel.Controls.Add(lbl);
            panel.Controls.Add(tb);
            return (panel, tb);
        }

        #endregion

        #region Tab 页面创建方法

        private Control CreateNetworkDiagTab()
        {
            var innerTabs = new TabControl { Dock = DockStyle.Fill };

            // --- Ping 子页 ---
            var pingPage = new TabPage("Ping");
            var pingPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var pingHost = LabeledTextBox("目标主机 / IP：", "127.0.0.1");
            var pingCount = LabeledTextBox("Ping 次数：", "4");
            var pingResp = new TextBox { Width = 900, Height = 300, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            var pingBtn = new Button { Text = "开始 Ping", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            pingBtn.Click += async (_, __) => await DoPing(pingHost.TextBox.Text, pingCount.TextBox.Text, pingResp);
            pingPanel.Controls.Add(pingHost.Panel);
            pingPanel.Controls.Add(pingCount.Panel);
            pingPanel.Controls.Add(pingBtn);
            pingPanel.Controls.Add(pingResp);
            pingPage.Controls.Add(pingPanel);
            innerTabs.TabPages.Add(pingPage);

            // --- DNS 解析子页 ---
            var dnsPage = new TabPage("DNS 解析");
            var dnsPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var dnsDomain = LabeledTextBox("域名：", "www.baidu.com");
            var dnsResp = new TextBox { Width = 900, Height = 350, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            var dnsBtn = new Button { Text = "解析 DNS", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            dnsBtn.Click += async (_, __) => await DoDnsLookup(dnsDomain.TextBox.Text, dnsResp);
            dnsPanel.Controls.Add(dnsDomain.Panel);
            dnsPanel.Controls.Add(dnsBtn);
            dnsPanel.Controls.Add(dnsResp);
            dnsPage.Controls.Add(dnsPanel);
            innerTabs.TabPages.Add(dnsPage);

            // --- 端口扫描子页 ---
            var scanPage = new TabPage("端口扫描");
            var scanPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var scanHost = LabeledTextBox("目标主机：", "127.0.0.1");
            var scanStart = LabeledTextBox("起始端口：", "1");
            var scanEnd = LabeledTextBox("结束端口：", "1024");
            var scanResp = new TextBox { Width = 900, Height = 300, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            var scanBtns = new FlowLayoutPanel { AutoSize = true };
            var scanBtn = new Button { Text = "开始扫描", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var scanStopBtn = new Button { Text = "停止扫描", AutoSize = true, Padding = new Padding(10, 5, 10, 5), Enabled = false };
            scanBtn.Click += async (_, __) =>
            {
                scanBtn.Enabled = false;
                scanStopBtn.Enabled = true;
                await DoPortScan(scanHost.TextBox.Text, scanStart.TextBox.Text, scanEnd.TextBox.Text, scanResp);
                scanBtn.Enabled = true;
                scanStopBtn.Enabled = false;
            };
            scanStopBtn.Click += (_, __) => { _portScanCts?.Cancel(); };
            scanBtns.Controls.Add(scanBtn);
            scanBtns.Controls.Add(scanStopBtn);
            scanPanel.Controls.Add(scanHost.Panel);
            scanPanel.Controls.Add(scanStart.Panel);
            scanPanel.Controls.Add(scanEnd.Panel);
            scanPanel.Controls.Add(scanBtns);
            scanPanel.Controls.Add(new Label { Text = "提示：范围不超过 1024 个端口，扫描会实时显示进度，可随时停止", Width = 900, ForeColor = System.Drawing.Color.Gray });
            scanPanel.Controls.Add(scanResp);
            scanPage.Controls.Add(scanPanel);
            innerTabs.TabPages.Add(scanPage);

            // --- SSL 证书检查子页 ---
            var sslPage = new TabPage("SSL 证书检查");
            var sslPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var sslHost = LabeledTextBox("目标主机：", "www.baidu.com");
            var sslPort = LabeledTextBox("端口：", "443");
            var sslResp = new TextBox { Width = 900, Height = 350, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            var sslBtn = new Button { Text = "检查证书", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            sslBtn.Click += async (_, __) => await DoSslCertCheck(sslHost.TextBox.Text, sslPort.TextBox.Text, sslResp);
            sslPanel.Controls.Add(sslHost.Panel);
            sslPanel.Controls.Add(sslPort.Panel);
            sslPanel.Controls.Add(sslBtn);
            sslPanel.Controls.Add(sslResp);
            sslPage.Controls.Add(sslPanel);
            innerTabs.TabPages.Add(sslPage);

            return innerTabs;
        }

        private Control CreateConnectivityTab()
        {
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10, 10, 10, 10) };

            var ip = LabeledTextBox("目标 IP / 主机：", "127.0.0.1");
            var port = LabeledTextBox("目标端口：", "3306");
            var btnIpTest = new Button { Text = "检测 IP:Port 可达性", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnIpTest.Click += async (_, __) => { var result = await TestIpPort(ip.TextBox.Text, port.TextBox.Text); };

            panel.Controls.Add(new Label { Text = "—— IP + Port 通断检测 ——", Width = 940 });
            panel.Controls.Add(ip.Panel);
            panel.Controls.Add(port.Panel);
            panel.Controls.Add(btnIpTest);

            panel.Controls.Add(new Label { Text = "—— 数据库连接检测 ——", Width = 940 });

            var dbServer = LabeledTextBox("数据库主机：", "localhost");
            var dbPort = LabeledTextBox("数据库端口：", "");
            var dbName = LabeledTextBox("数据库名 / 服务名：", "testdb");
            var dbUser = LabeledTextBox("用户名：", "");
            var dbPass = LabeledTextBox("密码：", "");
            var dbConnStr = LabeledTextBox("完整连接字符串：", "", multiline: true, height: 60);

            var btnTestSqlServer = new Button { Text = "测试 SQL Server", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnTestSqlServer.Click += async (_, __) =>
            {
                var cs = string.IsNullOrWhiteSpace(dbConnStr.TextBox.Text) ? BuildSqlServerConnectionString(dbServer.TextBox.Text, dbPort.TextBox.Text, dbName.TextBox.Text, dbUser.TextBox.Text, dbPass.TextBox.Text) : dbConnStr.TextBox.Text;
                await TestDbConnection("sqlserver", cs);
            };

            var btnTestOracle = new Button { Text = "测试 Oracle", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnTestOracle.Click += async (_, __) =>
            {
                var cs = string.IsNullOrWhiteSpace(dbConnStr.TextBox.Text) ? BuildOracleConnectionString(dbServer.TextBox.Text, dbPort.TextBox.Text, dbName.TextBox.Text, dbUser.TextBox.Text, dbPass.TextBox.Text) : dbConnStr.TextBox.Text;
                await TestDbConnection("oracle", cs);
            };

            var btnTestMySql = new Button { Text = "测试 MySQL", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnTestMySql.Click += async (_, __) =>
            {
                var cs = string.IsNullOrWhiteSpace(dbConnStr.TextBox.Text) ? BuildMySqlConnectionString(dbServer.TextBox.Text, dbPort.TextBox.Text, dbName.TextBox.Text, dbUser.TextBox.Text, dbPass.TextBox.Text) : dbConnStr.TextBox.Text;
                await TestDbConnection("mysql", cs);
            };

            var btnTestPgSql = new Button { Text = "测试 PostgreSQL", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnTestPgSql.Click += async (_, __) =>
            {
                var cs = string.IsNullOrWhiteSpace(dbConnStr.TextBox.Text) ? BuildNpgsqlConnectionString(dbServer.TextBox.Text, dbPort.TextBox.Text, dbName.TextBox.Text, dbUser.TextBox.Text, dbPass.TextBox.Text) : dbConnStr.TextBox.Text;
                await TestDbConnection("npgsql", cs);
            };

            panel.Controls.Add(dbServer.Panel);
            panel.Controls.Add(dbPort.Panel);
            panel.Controls.Add(dbName.Panel);
            panel.Controls.Add(dbUser.Panel);
            panel.Controls.Add(dbPass.Panel);
            panel.Controls.Add(dbConnStr.Panel);

            var btnPanel = new FlowLayoutPanel { AutoSize = true };
            btnPanel.Controls.Add(btnTestSqlServer);
            btnPanel.Controls.Add(btnTestOracle);
            btnPanel.Controls.Add(btnTestMySql);
            btnPanel.Controls.Add(btnTestPgSql);
            panel.Controls.Add(btnPanel);

            var dgvPanel = new Panel { Width = 940, Height = 240 };
            _dbResultGrid = new DataGridView
            {
                Left = 0, Top = 0, Width = 940, Height = 240,
                ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, ColumnCount = 6
            };
            _dbResultGrid.Columns[0].Name = "时间";
            _dbResultGrid.Columns[1].Name = "数据库类型";
            _dbResultGrid.Columns[2].Name = "目标";
            _dbResultGrid.Columns[3].Name = "结果";
            _dbResultGrid.Columns[4].Name = "耗时(ms)";
            _dbResultGrid.Columns[5].Name = "消息/异常";
            dgvPanel.Controls.Add(_dbResultGrid);
            panel.Controls.Add(new Label { Text = "—— 数据库检测结果（表格） ——", Width = 940 });
            panel.Controls.Add(dgvPanel);

            return panel;
        }

        private Control CreateRestTab()
        {
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10, 10, 10, 10) };
            var url = LabeledTextBox("接口 URL（POST）：", "https://postman-echo.com/post");
            var json = LabeledTextBox("JSON 请求体：", "{\"a\":1}", multiline: true, height: 120);
            var btn = new Button { Text = "POST JSON 并显示响应", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btn.Click += async (_, __) => await DoRestPost(url.TextBox.Text, json.TextBox.Text);

            var respPanel = new Panel { Width = 940, Height = 300 };
            var respLabel = new Label { Text = "响应（Body）:", Left = 0, Top = 4, Width = 200 };
            _restResponseBox = new TextBox { Left = 210, Top = 0, Width = 700, Height = 290, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            respPanel.Controls.Add(respLabel);
            respPanel.Controls.Add(_restResponseBox);
            panel.Controls.Add(url.Panel);
            panel.Controls.Add(json.Panel);
            panel.Controls.Add(btn);
            panel.Controls.Add(respPanel);
            return panel;
        }

        private Control CreateHttpTab()
        {
            var page = new TabPage("HTTP/HTTPS");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10, 10, 10, 10) };
            var urlBox = LabeledTextBox("URL（http/https）：", "https://postman-echo.com/get");

            var methodPanel = new FlowLayoutPanel { AutoSize = true, Width = 940 };
            var btnGet = new Button { Text = "GET", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var btnPost = new Button { Text = "POST", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var btnPut = new Button { Text = "PUT", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var btnDelete = new Button { Text = "DELETE", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var btnPatch = new Button { Text = "PATCH", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            methodPanel.Controls.Add(btnGet);
            methodPanel.Controls.Add(btnPost);
            methodPanel.Controls.Add(btnPut);
            methodPanel.Controls.Add(btnDelete);
            methodPanel.Controls.Add(btnPatch);

            var reqBody = LabeledTextBox("请求体：", "{\"hello\":\"world\"}", multiline: true, height: 80);
            var headersBox = LabeledTextBox("自定义 Headers（每行一个 Key: Value）：", "", multiline: true, height: 60, labelMultiline: true);

            var authTypePanel = new Panel { Width = 940, Height = 34 };
            var authLabel = new Label { Left = 0, Top = 2, Width = 200, Height = 26, Text = "认证方式：" };
            var authCombo = new ComboBox { Left = 210, Top = 2, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            authCombo.Items.AddRange(new[] { "None", "Basic", "Bearer" });
            authCombo.SelectedIndex = 0;
            authTypePanel.Controls.Add(authLabel);
            authTypePanel.Controls.Add(authCombo);

            var authUser = LabeledTextBox("Basic 用户名：", "");
            var authPass = LabeledTextBox("Basic 密码：", "");
            var authToken = LabeledTextBox("Bearer Token：", "");
            authCombo.SelectedIndexChanged += (_, __) =>
            {
                authUser.Panel.Visible = authCombo.SelectedItem?.ToString() == "Basic";
                authPass.Panel.Visible = authCombo.SelectedItem?.ToString() == "Basic";
                authToken.Panel.Visible = authCombo.SelectedItem?.ToString() == "Bearer";
            };
            authUser.Panel.Visible = false;
            authPass.Panel.Visible = false;
            authToken.Panel.Visible = false;

            var headerPanel = new Panel { Width = 940, Height = 120 };
            var headerLabel = new Label { Text = "响应 Headers:", Left = 0, Top = 4, Width = 200 };
            _httpHeaderBox = new TextBox { Left = 210, Top = 0, Width = 700, Height = 110, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 9) };
            headerPanel.Controls.Add(headerLabel);
            headerPanel.Controls.Add(_httpHeaderBox);

            var respPanel = new Panel { Width = 940, Height = 300 };
            var respLabel = new Label { Text = "响应 Body:", Left = 0, Top = 4, Width = 200 };
            _httpResponseBox = new TextBox { Left = 210, Top = 0, Width = 700, Height = 290, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            respPanel.Controls.Add(respLabel);
            respPanel.Controls.Add(_httpResponseBox);

            btnGet.Click += async (_, __) => await DoHttpRequest("GET", urlBox.TextBox.Text, null, headersBox.TextBox.Text, authCombo.SelectedItem?.ToString(), authUser.TextBox.Text, authPass.TextBox.Text, authToken.TextBox.Text, _httpResponseBox, _httpHeaderBox);
            btnPost.Click += async (_, __) => await DoHttpRequest("POST", urlBox.TextBox.Text, reqBody.TextBox.Text, headersBox.TextBox.Text, authCombo.SelectedItem?.ToString(), authUser.TextBox.Text, authPass.TextBox.Text, authToken.TextBox.Text, _httpResponseBox, _httpHeaderBox);
            btnPut.Click += async (_, __) => await DoHttpRequest("PUT", urlBox.TextBox.Text, reqBody.TextBox.Text, headersBox.TextBox.Text, authCombo.SelectedItem?.ToString(), authUser.TextBox.Text, authPass.TextBox.Text, authToken.TextBox.Text, _httpResponseBox, _httpHeaderBox);
            btnDelete.Click += async (_, __) => await DoHttpRequest("DELETE", urlBox.TextBox.Text, reqBody.TextBox.Text, headersBox.TextBox.Text, authCombo.SelectedItem?.ToString(), authUser.TextBox.Text, authPass.TextBox.Text, authToken.TextBox.Text, _httpResponseBox, _httpHeaderBox);
            btnPatch.Click += async (_, __) => await DoHttpRequest("PATCH", urlBox.TextBox.Text, reqBody.TextBox.Text, headersBox.TextBox.Text, authCombo.SelectedItem?.ToString(), authUser.TextBox.Text, authPass.TextBox.Text, authToken.TextBox.Text, _httpResponseBox, _httpHeaderBox);

            panel.Controls.Add(urlBox.Panel);
            panel.Controls.Add(new Label { Text = "—— HTTP 方法 ——", Width = 940 });
            panel.Controls.Add(methodPanel);
            panel.Controls.Add(reqBody.Panel);
            panel.Controls.Add(headersBox.Panel);
            panel.Controls.Add(new Label { Text = "—— 认证（可选）——", Width = 940 });
            panel.Controls.Add(authTypePanel);
            panel.Controls.Add(authUser.Panel);
            panel.Controls.Add(authPass.Panel);
            panel.Controls.Add(authToken.Panel);
            panel.Controls.Add(new Label { Text = "—— 响应 ——", Width = 940 });
            panel.Controls.Add(headerPanel);
            panel.Controls.Add(respPanel);
            page.Controls.Add(panel);
            return page;
        }

        private Control CreateWebSocketTab()
        {
            var page = new TabPage("WebSocket");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10, 10, 10, 10) };
            var urlBox = LabeledTextBox("WebSocket 地址：", "wss://echo.websocket.org");
            var btnConnect = new Button { Text = "连接并发送 'hello'", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnConnect.Click += async (_, __) => await DoWebSocketEcho(urlBox.TextBox.Text);
            panel.Controls.Add(urlBox.Panel);
            panel.Controls.Add(btnConnect);
            page.Controls.Add(panel);
            return page;
        }

        private Control CreateWebServiceTab()
        {
            var page = new TabPage("WebService (SOAP)");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10, 10, 10, 10) };
            var url = LabeledTextBox("服务 URL：", "https://www.example.com/Service.svc");
            var soapAction = LabeledTextBox("SOAPAction（可选）：", "");
            var reqXml = LabeledTextBox("请求 XML：", "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\">\n  <soapenv:Body>\n    <!-- your request here -->\n  </soapenv:Body>\n</soapenv:Envelope>", multiline: true, height: 200);
            var btn = new Button { Text = "调用 WebService (SOAP)", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var respPanel = new Panel { Width = 940, Height = 320 };
            var respLabel = new Label { Text = "响应（Body）:", Left = 0, Top = 4, Width = 200 };
            _wsResponseBox = new TextBox { Left = 210, Top = 0, Width = 700, Height = 300, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            respPanel.Controls.Add(respLabel);
            respPanel.Controls.Add(_wsResponseBox);
            btn.Click += async (_, __) => await DoWebServiceTest(url.TextBox.Text, soapAction.TextBox.Text, reqXml.TextBox.Text);
            panel.Controls.Add(url.Panel);
            panel.Controls.Add(soapAction.Panel);
            panel.Controls.Add(reqXml.Panel);
            panel.Controls.Add(new Label { Text = "注意：如果使用 HTTPS，可能需要受信任的证书；某些服务要求额外认证（如 Basic、WS-Security）。", Width = 900,Height = 120, ForeColor = System.Drawing.Color.DarkRed, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Italic) });
            panel.Controls.Add(btn);
            panel.Controls.Add(respPanel);
            page.Controls.Add(panel);
            return page;
        }

        private Control CreateFtpSftpTab()
        {
            var page = new TabPage("FTP / SFTP / SSH");
            var innerTabs = new TabControl { Dock = DockStyle.Fill };

            // === FTP/SFTP 子页 ===
            var ftpPage = new TabPage("FTP / SFTP");
            var ftpPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var host = LabeledTextBox("主机（Host）：", "ftp.example.com");
            var port = LabeledTextBox("端口（Port，可留空使用默认）：", "", labelMultiline: true);
            var user = LabeledTextBox("用户名：", "anonymous");
            var pass = LabeledTextBox("密码：", "");
            var keyPath = LabeledTextBox("私钥路径：", "", multiline: false);

            var btnFtpList = new Button { Text = "FTP：列出目录", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnFtpList.Click += async (_, __) => await DoFtpList(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text);
            var btnSftpList = new Button { Text = "SFTP：列出根目录", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnSftpList.Click += async (_, __) => await DoSftpList_Strong(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text, keyPath.TextBox.Text);

            var ftpOpLabel = new Label { Text = "—— FTP 文件操作 ——", Width = 900, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Bold) };
            var remotePath = LabeledTextBox("远程路径：", "/upload/test.txt");
            var localPath = LabeledTextBox("本地路径：", @"C:\temp\test.txt", multiline: false);
            var btnUpload = new Button { Text = "上传", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var btnDownload = new Button { Text = "下载", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var btnDelete = new Button { Text = "删除", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var btnMkdir = new Button { Text = "创建目录", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnUpload.Click += async (_, __) => await DoFtpUpload(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text, localPath.TextBox.Text, remotePath.TextBox.Text);
            btnDownload.Click += async (_, __) => await DoFtpDownload(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text, remotePath.TextBox.Text, localPath.TextBox.Text);
            btnDelete.Click += async (_, __) => await DoFtpDelete(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text, remotePath.TextBox.Text);
            btnMkdir.Click += async (_, __) => await DoFtpMkdir(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text, remotePath.TextBox.Text);
            var ftpOpBtns = new FlowLayoutPanel { AutoSize = true };
            ftpOpBtns.Controls.Add(btnUpload);
            ftpOpBtns.Controls.Add(btnDownload);
            ftpOpBtns.Controls.Add(btnDelete);
            ftpOpBtns.Controls.Add(btnMkdir);
            _ftpResponseBox = new TextBox { Width = 900, Height = 100, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };

            ftpPanel.Controls.Add(host.Panel);
            ftpPanel.Controls.Add(port.Panel);
            ftpPanel.Controls.Add(user.Panel);
            ftpPanel.Controls.Add(pass.Panel);
            ftpPanel.Controls.Add(keyPath.Panel);
            ftpPanel.Controls.Add(btnFtpList);
            ftpPanel.Controls.Add(btnSftpList);
            ftpPanel.Controls.Add(new Label { Text = "注意：FTP 可支持匿名访问，但很多服务器禁止匿名；SFTP 使用 SSH 身份验证。", Width = 900, ForeColor = System.Drawing.Color.DarkRed, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Italic) });
            ftpPanel.Controls.Add(ftpOpLabel);
            ftpPanel.Controls.Add(remotePath.Panel);
            ftpPanel.Controls.Add(localPath.Panel);
            ftpPanel.Controls.Add(ftpOpBtns);
            ftpPanel.Controls.Add(_ftpResponseBox);
            ftpPage.Controls.Add(ftpPanel);
            innerTabs.TabPages.Add(ftpPage);

            // === SSH 命令子页 ===
            var sshPage = new TabPage("SSH 命令");
            var sshPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var sshHost = LabeledTextBox("SSH 主机：", "192.168.1.100");
            var sshPort = LabeledTextBox("端口：", "22");
            var sshUser = LabeledTextBox("用户名：", "root");
            var sshPass = LabeledTextBox("密码：", "");
            var sshCmd = LabeledTextBox("命令：", "uname -a", multiline: true, height: 60);
            _sshResponseBox = new TextBox { Width = 900, Height = 350, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            var sshBtn = new Button { Text = "执行命令", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            sshBtn.Click += async (_, __) => await DoSshCommand(sshHost.TextBox.Text, sshPort.TextBox.Text, sshUser.TextBox.Text, sshPass.TextBox.Text, sshCmd.TextBox.Text, _sshResponseBox);
            sshPanel.Controls.Add(sshHost.Panel);
            sshPanel.Controls.Add(sshPort.Panel);
            sshPanel.Controls.Add(sshUser.Panel);
            sshPanel.Controls.Add(sshPass.Panel);
            sshPanel.Controls.Add(sshCmd.Panel);
            sshPanel.Controls.Add(sshBtn);
            sshPanel.Controls.Add(new Label { Text = "提示：使用 Renci.SshNet 库，支持密码认证。命令在远程服务器上执行。", Width = 900, ForeColor = System.Drawing.Color.Gray });
            sshPanel.Controls.Add(_sshResponseBox);
            sshPage.Controls.Add(sshPanel);
            innerTabs.TabPages.Add(sshPage);

            return innerTabs;
        }

        private Control CreateSmtpTab()
        {
            var page = new TabPage("SMTP");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10, 10, 10, 10) };
            var host = LabeledTextBox("SMTP 主机：", "smtp.example.com");
            var port = LabeledTextBox("端口：", "25");
            var user = LabeledTextBox("用户名：", "");
            var pass = LabeledTextBox("密码：", "");
            var from = LabeledTextBox("发件人（From）：", "sender@example.com");
            var to = LabeledTextBox("收件人（To）：", "recipient@example.com");
            var subj = LabeledTextBox("主题：", "测试邮件");
            var body = LabeledTextBox("正文：", "来自协议调试器的测试邮件", multiline: true, height: 100);
            var sslPanel = new Panel { Width = 940, Height = 34 };
            var sslCheck = new CheckBox { Left = 210, Top = 2, Width = 200, Checked = false, Text = "使用 SSL/TLS" };
            var sslLabel = new Label { Left = 0, Top = 2, Width = 200, Height = 26, Text = "安全连接：" };
            sslPanel.Controls.Add(sslLabel);
            sslPanel.Controls.Add(sslCheck);
            panel.Controls.Add(new Label { Text = "注意：大多数 SMTP 服务器要求身份验证和 TLS/SSL。常见端口：25 (可 STARTTLS)、465 (SSL)、587 (提交)。匿名发送通常被拒绝。", Width = 940, Height = 120, ForeColor = System.Drawing.Color.DarkRed, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Italic) });
            var btnSend = new Button { Text = "发送邮件（SmtpClient）", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnSend.Click += async (_, __) => await DoSendSmtp(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text, from.TextBox.Text, to.TextBox.Text, subj.TextBox.Text, body.TextBox.Text, sslCheck.Checked);
            panel.Controls.Add(host.Panel);
            panel.Controls.Add(port.Panel);
            panel.Controls.Add(user.Panel);
            panel.Controls.Add(pass.Panel);
            panel.Controls.Add(sslPanel);
            panel.Controls.Add(from.Panel);
            panel.Controls.Add(to.Panel);
            panel.Controls.Add(subj.Panel);
            panel.Controls.Add(body.Panel);
            panel.Controls.Add(btnSend);
            page.Controls.Add(panel);
            return page;
        }

        private Control CreateDeviceProtocolsTab()
        {
            var page = new TabPage("上位机协议测试");
            var innerTabs = new TabControl { Dock = DockStyle.Fill };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var topPanel = new Panel { Dock = DockStyle.Fill };
            topPanel.Controls.Add(new Label { Text = "注意：设备协议通常在受控网络环境中，需要设备侧账号、访问控制或特殊路由。防火墙、端口和协议版本（例如 S7、Modbus 变体）会影响可用性。", Width = 900, Height = 120, Dock = DockStyle.Fill, ForeColor = System.Drawing.Color.DarkRed, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Italic), Padding = new Padding(6) });
            layout.Controls.Add(topPanel, 0, 0);
            innerTabs.Dock = DockStyle.Fill;
            layout.Controls.Add(innerTabs, 0, 1);

            // S7
            var s7Page = new TabPage("S7");
            var s7Panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var s7Host = LabeledTextBox("目标主机：", "192.168.1.100");
            var s7Port = LabeledTextBox("端口：", "102");
            var s7Address = LabeledTextBox("读写地址：", "DB1.DBW0");
            var s7WriteVal = LabeledTextBox("写入值：", "123");
            var s7ReadBtn = new Button { Text = "读 S7 点位", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var s7WriteBtn = new Button { Text = "写 S7 点位", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var s7Resp = new TextBox { Left = 210, Top = 0, Width = 680, Height = 120, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            s7ReadBtn.Click += async (_, __) => { _deviceProtoResponseBox = s7Resp; await TestS7ReadWrite(s7Host.TextBox.Text, s7Port.TextBox.Text, s7Address.TextBox.Text, null); };
            s7WriteBtn.Click += async (_, __) => { _deviceProtoResponseBox = s7Resp; await TestS7ReadWrite(s7Host.TextBox.Text, s7Port.TextBox.Text, s7Address.TextBox.Text, s7WriteVal.TextBox.Text); };
            s7Panel.Controls.Add(s7Host.Panel);
            s7Panel.Controls.Add(s7Port.Panel);
            s7Panel.Controls.Add(s7Address.Panel);
            s7Panel.Controls.Add(s7WriteVal.Panel);
            s7Panel.Controls.Add(s7ReadBtn);
            s7Panel.Controls.Add(s7WriteBtn);
            s7Panel.Controls.Add(s7Resp);
            s7Page.Controls.Add(s7Panel);
            innerTabs.TabPages.Add(s7Page);

            // Modbus
            var modbusPage = new TabPage("Modbus TCP");
            var modbusPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var mbHost = LabeledTextBox("目标主机：", "192.168.1.100");
            var mbPort = LabeledTextBox("端口：", "502");
            var mbUnit = LabeledTextBox("UnitId：", "1");
            var mbStart = LabeledTextBox("起始地址：", "0");
            var mbQty = LabeledTextBox("读取数量：", "1");
            var mbBtn = new Button { Text = "读取保持寄存器 (FC03)", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var mbResp = new TextBox { Left = 210, Top = 0, Width = 680, Height = 120, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            mbBtn.Click += async (_, __) => { _deviceProtoResponseBox = mbResp; await TestModbusTcp(mbHost.TextBox.Text, mbPort.TextBox.Text, mbUnit.TextBox.Text, mbStart.TextBox.Text, mbQty.TextBox.Text); };
            modbusPanel.Controls.Add(mbHost.Panel);
            modbusPanel.Controls.Add(mbPort.Panel);
            modbusPanel.Controls.Add(mbUnit.Panel);
            modbusPanel.Controls.Add(mbStart.Panel);
            modbusPanel.Controls.Add(mbQty.Panel);
            modbusPanel.Controls.Add(mbBtn);
            modbusPanel.Controls.Add(new Label { Text = "—— Modbus 写入 ——", Width = 680, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Bold) });
            var mbFcPanel = new Panel { Width = 680, Height = 34 };
            var mbFcLabel = new Label { Left = 0, Top = 2, Width = 200, Height = 26, Text = "功能码：" };
            var mbFcCombo = new ComboBox { Left = 210, Top = 2, Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
            mbFcCombo.Items.AddRange(new[] { "FC06 - 写单个寄存器", "FC16 - 写多个寄存器", "FC05 - 写单个线圈" });
            mbFcCombo.SelectedIndex = 0;
            mbFcPanel.Controls.Add(mbFcLabel);
            mbFcPanel.Controls.Add(mbFcCombo);
            modbusPanel.Controls.Add(mbFcPanel);
            var mbWriteAddr = LabeledTextBox("写入起始地址：", "0");
            var mbWriteVal = LabeledTextBox("写入值：", "123");
            var mbWriteBtn = new Button { Text = "执行写入", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            mbWriteBtn.Click += async (_, __) => { _deviceProtoResponseBox = mbResp; var fc = mbFcCombo.SelectedIndex switch { 0 => "6", 1 => "16", 2 => "5", _ => "6" }; await TestModbusWrite(mbHost.TextBox.Text, mbPort.TextBox.Text, mbUnit.TextBox.Text, fc, mbWriteAddr.TextBox.Text, mbWriteVal.TextBox.Text, mbResp); };
            modbusPanel.Controls.Add(mbWriteAddr.Panel);
            modbusPanel.Controls.Add(mbWriteVal.Panel);
            modbusPanel.Controls.Add(mbWriteBtn);
            modbusPanel.Controls.Add(mbResp);
            modbusPage.Controls.Add(modbusPanel);
            innerTabs.TabPages.Add(modbusPage);

            // OPC
            var opcPage = new TabPage("OPC");
            var opcPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var uaEndpoint = LabeledTextBox("Endpoint URL:", "opc.tcp://localhost:4840");
            var uaNode = LabeledTextBox("NodeId:", "ns=2;s=Demo.Static.Scalar.Int32");
            var uaReadBtn = new Button { Text = "OPC UA: 读节点", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var uaWriteBtn = new Button { Text = "OPC UA: 写节点", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var uaWriteVal = LabeledTextBox("写入值：", "123");
            var uaResp = new TextBox { Left = 210, Top = 0, Width = 680, Height = 120, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            uaReadBtn.Click += async (_, __) => { _opcUaRespBox = uaResp; _deviceProtoResponseBox = uaResp; await DoOpcUaRead(uaEndpoint.TextBox.Text, uaNode.TextBox.Text); };
            uaWriteBtn.Click += async (_, __) => { _opcUaRespBox = uaResp; _deviceProtoResponseBox = uaResp; await DoOpcUaWrite(uaEndpoint.TextBox.Text, uaNode.TextBox.Text, uaWriteVal.TextBox.Text); };
            opcPanel.Controls.Add(uaEndpoint.Panel);
            opcPanel.Controls.Add(uaNode.Panel);
            opcPanel.Controls.Add(uaWriteVal.Panel);
            opcPanel.Controls.Add(uaReadBtn);
            opcPanel.Controls.Add(uaWriteBtn);
            opcPanel.Controls.Add(uaResp);
            opcPanel.Controls.Add(new Label { Text = "OPC DA（本机 COM）", Width = 900 });
            var daProg = LabeledTextBox("ProgID:", "OPC.Automation.1");
            var daItemId = LabeledTextBox("ItemID:", "");
            var daReadBtn = new Button { Text = "OPC DA: 读项", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var daWriteBtn = new Button { Text = "OPC DA: 写项", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var daWriteVal = LabeledTextBox("写入值：", "123");
            var daResp = new TextBox { Left = 210, Top = 0, Width = 680, Height = 120, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            daReadBtn.Click += async (_, __) => { _deviceProtoResponseBox = daResp; await DoOpcDaReadWrite("", daProg.TextBox.Text, daItemId.TextBox.Text, null); };
            daWriteBtn.Click += async (_, __) => { _deviceProtoResponseBox = daResp; await DoOpcDaReadWrite("", daProg.TextBox.Text, daItemId.TextBox.Text, daWriteVal.TextBox.Text); };
            opcPanel.Controls.Add(daProg.Panel);
            opcPanel.Controls.Add(daItemId.Panel);
            opcPanel.Controls.Add(daWriteVal.Panel);
            opcPanel.Controls.Add(daReadBtn);
            opcPanel.Controls.Add(daWriteBtn);
            opcPanel.Controls.Add(daResp);
            opcPage.Controls.Add(opcPanel);
            innerTabs.TabPages.Add(opcPage);

            // Raw TCP
            var rawPage = new TabPage("Raw TCP");
            var rawPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var rawHost = LabeledTextBox("目标主机：", "192.168.1.100");
            var rawPort = LabeledTextBox("端口：", "502");
            var rawPayload = LabeledTextBox("发送内容：", "Hello", multiline: true, height: 80);
            var rawBtn = new Button { Text = "发送并接收", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            var rawResp = new TextBox { Left = 210, Top = 0, Width = 680, Height = 120, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            rawBtn.Click += async (_, __) => { _deviceProtoResponseBox = rawResp; await TestRawTcp(rawHost.TextBox.Text, rawPort.TextBox.Text, rawPayload.TextBox.Text); };
            rawPanel.Controls.Add(rawHost.Panel);
            rawPanel.Controls.Add(rawPort.Panel);
            rawPanel.Controls.Add(rawPayload.Panel);
            rawPanel.Controls.Add(rawBtn);
            rawPanel.Controls.Add(rawResp);
            rawPage.Controls.Add(rawPanel);
            innerTabs.TabPages.Add(rawPage);

            page.Controls.Add(layout);
            return page;
        }

        private Control CreateMailReceiveTab()
        {
            var page = new TabPage("收取邮件 (MailKit)");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10, 10, 10, 10) };
            var host = LabeledTextBox("POP3 主机：", "pop.example.com");
            var port = LabeledTextBox("端口：", "995");
            var user = LabeledTextBox("用户名：", "");
            var pass = LabeledTextBox("密码：", "");
            var sslPanel = new Panel { Width = 940, Height = 34 };
            var sslCheck = new CheckBox { Left = 210, Top = 2, Width = 200, Checked = true, Text = "使用 SSL/TLS" };
            sslPanel.Controls.Add(new Label { Left = 0, Top = 2, Width = 200, Height = 26, Text = "安全连接：" });
            sslPanel.Controls.Add(sslCheck);
            var pop3Btn = new Button { Text = "POP3：列出前 5 封（MailKit）", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            pop3Btn.Click += async (_, __) => await DoPop3MailKit(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text, sslCheck.Checked);
            var imapHost = LabeledTextBox("IMAP 主机：", "imap.example.com");
            var imapPort = LabeledTextBox("IMAP 端口：", "993");
            var imapUser = LabeledTextBox("IMAP 用户名：", "");
            var imapPass = LabeledTextBox("IMAP 密码：", "");
            var imapSslPanel = new Panel { Width = 940, Height = 34 };
            var imapSslCheck = new CheckBox { Left = 210, Top = 2, Width = 200, Checked = true, Text = "使用 SSL/TLS" };
            imapSslPanel.Controls.Add(new Label { Left = 0, Top = 2, Width = 200, Height = 26, Text = "IMAP 安全：" });
            imapSslPanel.Controls.Add(imapSslCheck);
            var imapBtn = new Button { Text = "IMAP：列出收件箱前 10 封（MailKit）", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            imapBtn.Click += async (_, __) => await DoImapMailKit(imapHost.TextBox.Text, imapPort.TextBox.Text, imapUser.TextBox.Text, imapPass.TextBox.Text, imapSslCheck.Checked);
            panel.Controls.Add(host.Panel);
            panel.Controls.Add(port.Panel);
            panel.Controls.Add(user.Panel);
            panel.Controls.Add(pass.Panel);
            panel.Controls.Add(sslPanel);
            panel.Controls.Add(pop3Btn);
            panel.Controls.Add(new Label { Text = "注意：POP3/IMAP 服务通常要求 TLS/SSL 和账号认证；部分大厂使用 OAuth2 登录（无法直接使用密码），请根据服务文档配置。", Width = 940, ForeColor = System.Drawing.Color.DarkRed, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Italic) });
            panel.Controls.Add(new Label { Text = "—— IMAP（可选） ——", Width = 940 });
            panel.Controls.Add(imapHost.Panel);
            panel.Controls.Add(imapPort.Panel);
            panel.Controls.Add(imapUser.Panel);
            panel.Controls.Add(imapPass.Panel);
            panel.Controls.Add(imapSslPanel);
            panel.Controls.Add(imapBtn);
            page.Controls.Add(panel);
            return page;
        }

        private Control CreateLogTab()
        {
            var page = new TabPage("日志");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10, 10, 10, 10) };
            _logBox = new TextBox { Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10), Width = 940, Height = 600 };
            var btns = new FlowLayoutPanel { AutoSize = true };
            var btnClear = new Button { Text = "清除日志", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnClear.Click += (_, __) => { lock (_logLock) { _logLines.Clear(); _logBox.Clear(); } };
            var btnExport = new Button { Text = "导出日志到文件", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnExport.Click += (_, __) => { try { var dlg = new SaveFileDialog { Filter = "文本文件|*.txt", FileName = $"protocol-debug-log-{DateTime.Now:yyyyMMddHHmmss}.txt" }; if (dlg.ShowDialog() == DialogResult.OK) { File.WriteAllLines(dlg.FileName, _logLines); MessageBox.Show("导出成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); } } catch (Exception ex) { MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); } };
            var btnOpenFolder = new Button { Text = "打开日志文件夹", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnOpenFolder.Click += (_, __) => OpenLogFolder();
            var btnExportToday = new Button { Text = "导出今日日志", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
            btnExportToday.Click += (_, __) => ExportTodayLog();
            btns.Controls.Add(btnClear);
            btns.Controls.Add(btnExport);
            btns.Controls.Add(btnOpenFolder);
            btns.Controls.Add(btnExportToday);
            panel.Controls.Add(btns);
            panel.Controls.Add(_logBox);
            page.Controls.Add(panel);
            return page;
        }

        private Control CreateHelpTab()
        {
            var page = new TabPage("帮助");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10, 10, 10, 10) };
            var helpLabel = new Label { Text = "使用教程", Font = new System.Drawing.Font("Microsoft Sans Serif", 11, System.Drawing.FontStyle.Bold), Width = 940, Margin = new Padding(0, 10, 0, 5) };
            var helpText = @"快速入门：

1) 网络诊断
 - Ping：输入目标主机和次数，点击""开始 Ping""查看延迟和丢包率。
 - DNS 解析：输入域名，查看 A/AAAA/CNAME 等记录。
 - 端口扫描：输入主机和端口范围（最多 1024 个），实时显示扫描进度，可随时停止。
 - SSL 证书检查：输入主机和端口，查看证书详情、有效期和证书链。

2) 连通性检测
 - IP+Port：填写目标主机和端口，点击""检测 IP:Port 可达性""测试 TCP 连通性。
 - 数据库：填写主机/端口/用户名/密码（或直接填写连接字符串），支持 SQL Server、Oracle、MySQL、PostgreSQL。

3) REST 接口测试
 - 填写 URL 和 JSON 请求体，点击""POST JSON""发送请求并查看响应。

4) HTTP / HTTPS
 - 支持 GET/POST/PUT/DELETE/PATCH 方法。
 - 可自定义 Headers（每行一个 Key: Value）。
 - 支持 Basic 和 Bearer 认证。
 - 分别显示响应 Headers 和 Body。

5) WebSocket
 - 填写 WebSocket 地址（ws:// 或 wss://），点击连接并发送测试消息。

6) WebService (SOAP)
 - 填写服务 URL、SOAPAction 和 XML 请求体，发送 SOAP 请求。

7) FTP / SFTP / SSH
 - FTP：支持列目录、上传、下载、删除、创建目录。
 - SFTP：支持密码和私钥认证，列出根目录。
 - SSH：填写主机/用户名/密码和命令，远程执行并查看 stdout/stderr。

8) SMTP 发送邮件
 - 填写 SMTP 主机、端口、用户名、密码、发件人、收件人、主题和正文。
 - 可选 SSL/TLS 加密连接。

9) 设备协议测试
 - S7：西门子 PLC 读写（DB 地址格式如 DB1.DBW0）。
 - Modbus TCP：FC03 读寄存器、FC05 写线圈、FC06 写单寄存器、FC16 写多寄存器。
 - OPC UA：通过 Endpoint URL 和 NodeId 读写节点。
 - OPC DA：通过 ProgID 进行本机 COM 读写。
 - Raw TCP：向指定主机端口发送原始数据并接收响应。

10) 收取邮件（MailKit）
 - POP3 / IMAP：填写服务器信息，可选 SSL/TLS，列出最近邮件摘要。

11) 日志
 - 所有操作自动记录到日志页和本地日志文件（按日期分割）。
 - 支持清除、导出、打开日志文件夹。

12) 帮助 / 更新日志
 - 本页为使用教程，""更新日志""页查看版本变更记录。

提示：
 - 左侧导航栏点击切换功能页。
 - 网络操作均为异步执行，不会阻塞界面。
 - 日志文件保存在程序目录下的 Log 文件夹中。";
            var helpBox = new TextBox { Left = 0, Top = 0, Width = 940,  Height = 700, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, Font = new System.Drawing.Font("Consolas", 10), Text = helpText };
            panel.Controls.Add(helpLabel);
            panel.Controls.Add(helpBox);
            page.Controls.Add(panel);
            return page;
        }

        private Control CreateUpdateLogTab()
        {
            var page = new TabPage("更新日志");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10, 10, 10, 10) };
            var updateLogText = @"版本 1.004 - 2026年7月
• 新增网络诊断功能（Ping、DNS、端口扫描、SSL 证书检查）
• HTTP 支持全方法（GET/POST/PUT/DELETE/PATCH）+ 自定义 Headers + 认证
• FTP 增加上传/下载/删除/创建目录功能
• 新增 SSH 远程命令执行
• Modbus 增加写操作（FC05/FC06/FC16）
• 新增 PostgreSQL 数据库连接支持
• 界面改为左侧导航栏布局

版本 1.003 - 2025年10月
• 添加说明
• 帮助页面

版本 1.002 - 2025年10月
• 修复SMTP测试默认不使用SSL的问题

版本 1.001 - 2025年10月
• 初始版本发布
• 支持HTTP/HTTPS协议测试
• 支持WebSocket连接测试
• 支持WebService (SOAP)调用
• 支持FTP/SFTP文件操作
• 支持SMTP邮件发送
• 支持REST API测试
• 支持数据库连通性检测
• 支持上位机协议测试(S7, Modbus, OPC)
• 支持邮件收取(POP3/IMAP)
• 集成日志记录功能";
            var updateLogLabel = new Label { Text = "更新日志", Font = new System.Drawing.Font("Microsoft Sans Serif", 11, System.Drawing.FontStyle.Bold), Width = 940, Margin = new Padding(0, 10, 0, 5) };
            var updateLogBox = new TextBox { Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 11), Width = 940, Height = 700, Text = updateLogText };
            panel.Controls.Add(updateLogLabel);
            panel.Controls.Add(updateLogBox);
            page.Controls.Add(panel);
            return page;
        }

        #endregion

        /// <summary>
        /// 确保日志目录存在（静默失败以避免影响主流程）。
        /// </summary>
        private void EnsureLogDirectory()
        {
            try
            {
                if (!Directory.Exists(_logDir)) Directory.CreateDirectory(_logDir);
            }
            catch { /* 忽略创建目录异常，日志写入将尝试失败处理 */ }
        }

        /// <summary>
        /// 将单行日志写入当天的文件，此方法只处理文件写入（不更新 UI）。
        /// 使用 _logLock 同步，防止并发写入冲突。
        /// </summary>
        /// <param name="line">要写入的日志行（不包含换行符）</param>
        private void WriteToDailyLog(string line)
        {
            try
            {
                lock (_logLock)
                {
                    EnsureLogDirectory();
                    var file = Path.Combine(_logDir, DateTime.Now.ToString("yyyyMMdd") + ".log");
                    File.AppendAllText(file, line + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch { /* 日志写入失败不抛出，避免影响主流程 */ }
        }

        /// <summary>
        /// 添加一条详细日志块（包含请求/响应/异常等），并尝试写入 UI 和当日日志文件。
        /// 该方法用于记录请求-响应对，便于故障排查。
        /// </summary>
        /// <param name="title">日志块标题（如 "HTTP GET Request"）</param>
        /// <param name="requestInfo">请求摘要信息（例如 URL、目标地址）</param>
        /// <param name="requestBody">请求体（可能是 JSON 或 XML）</param>
        /// <param name="responseInfo">响应摘要（例如状态码）</param>
        /// <param name="responseBody">响应正文</param>
        /// <param name="ex">如果发生异常，则传入异常对象以记录堆栈</param>
        private void AppendDetailedLog(string title, string requestInfo = null, string requestBody = null, string responseInfo = null, string responseBody = null, Exception ex = null)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"----- {title} [{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] -----");
            if (!string.IsNullOrEmpty(requestInfo)) sb.AppendLine("Request: " + requestInfo);
            if (!string.IsNullOrEmpty(requestBody)) sb.AppendLine("RequestBody:\n" + requestBody);
            if (!string.IsNullOrEmpty(responseInfo)) sb.AppendLine("Response: " + responseInfo);
            if (!string.IsNullOrEmpty(responseBody)) sb.AppendLine("ResponseBody:\n" + responseBody);
            if (ex != null) sb.AppendLine("Exception: " + ex.ToString());
            sb.AppendLine("----- End -----");

            var block = sb.ToString();
            lock (_logLock) { _logLines.Add(block); }

            // 如果在非 UI 线程调用，使用 BeginInvoke 更新 UI，避免阻塞调用线程
            if (this.InvokeRequired)
                this.BeginInvoke(new Action(() => _logBox?.AppendText(block + Environment.NewLine)));
            else
                _logBox?.AppendText(block + Environment.NewLine);

            WriteToDailyLog(block);
        }

        /// <summary>
        /// 通用 HTTP 请求方法，支持 GET/POST/PUT/DELETE/PATCH，自定义 Headers 和认证。
        /// </summary>
        /// <param name="method">HTTP 方法（GET/POST/PUT/DELETE/PATCH）</param>
        /// <param name="url">目标 URL</param>
        /// <param name="body">请求体（POST/PUT/PATCH 时使用）</param>
        /// <param name="headersText">自定义 Headers（每行一个 Key: Value）</param>
        /// <param name="authType">认证类型（None/Basic/Bearer）</param>
        /// <param name="authUser">Basic 认证用户名</param>
        /// <param name="authPass">Basic 认证密码</param>
        /// <param name="authToken">Bearer Token</param>
        /// <param name="respBox">响应正文显示控件</param>
        /// <param name="headerBox">响应头显示控件</param>
        private async Task DoHttpRequest(string method, string url, string body, string headersText,
            string authType, string authUser, string authPass, string authToken,
            TextBox respBox, TextBox headerBox)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                MessageBox.Show($"{method}：URL 为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            AppendLog($"{method} -> {url}");
            var sw = Stopwatch.StartNew();
            try
            {
                var httpMethod = method switch
                {
                    "GET" => HttpMethod.Get,
                    "POST" => HttpMethod.Post,
                    "PUT" => HttpMethod.Put,
                    "DELETE" => HttpMethod.Delete,
                    "PATCH" => HttpMethod.Patch,
                    _ => HttpMethod.Get
                };

                using var req = new HttpRequestMessage(httpMethod, url);

                // 添加请求体（POST/PUT/PATCH）
                if (!string.IsNullOrWhiteSpace(body) && (method == "POST" || method == "PUT" || method == "PATCH"))
                {
                    req.Content = new StringContent(body, Encoding.UTF8, "application/json");
                }

                // 解析自定义 Headers
                if (!string.IsNullOrWhiteSpace(headersText))
                {
                    foreach (var line in headersText.Split('\n'))
                    {
                        var trimmed = line.Trim();
                        if (string.IsNullOrEmpty(trimmed)) continue;
                        var colonIdx = trimmed.IndexOf(':');
                        if (colonIdx > 0)
                        {
                            var key = trimmed.Substring(0, colonIdx).Trim();
                            var value = trimmed.Substring(colonIdx + 1).Trim();
                            // Content-Type 等头需要设置在 Content 上
                            if (key.Equals("Content-Type", StringComparison.OrdinalIgnoreCase) && req.Content != null)
                            {
                                req.Content.Headers.Clear();
                                req.Content.Headers.TryAddWithoutValidation(key, value);
                            }
                            else
                            {
                                req.Headers.TryAddWithoutValidation(key, value);
                            }
                        }
                    }
                }

                // 添加认证
                switch (authType?.ToLower())
                {
                    case "basic":
                        if (!string.IsNullOrWhiteSpace(authUser))
                        {
                            var credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{authUser}:{authPass}"));
                            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", credentials);
                        }
                        break;
                    case "bearer":
                        if (!string.IsNullOrWhiteSpace(authToken))
                        {
                            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken);
                        }
                        break;
                }

                AppendDetailedLog($"{method} Request", url, body);

                using var resp = await _httpClient.SendAsync(req);
                sw.Stop();
                var respBody = await resp.Content.ReadAsStringAsync();

                // 构建响应头字符串
                var respHeaders = new StringBuilder();
                respHeaders.AppendLine($"HTTP/{resp.Version} {(int)resp.StatusCode} {resp.ReasonPhrase}");
                foreach (var h in resp.Headers) respHeaders.AppendLine($"{h.Key}: {string.Join(", ", h.Value)}");
                foreach (var h in resp.Content.Headers) respHeaders.AppendLine($"{h.Key}: {string.Join(", ", h.Value)}");

                // 状态摘要
                string statusLine = $"状态: {(int)resp.StatusCode} {resp.ReasonPhrase}  耗时: {sw.ElapsedMilliseconds}ms  响应长度: {respBody?.Length ?? 0}";

                if (respBox != null) respBox.Text = respBody ?? "";
                if (headerBox != null) headerBox.Text = respHeaders.ToString();

                AppendLog(statusLine);
                AppendDetailedLog($"{method} Response", url, null, statusLine, respBody);
            }
            catch (Exception ex)
            {
                sw.Stop();
                string errorMsg = $"{method} 错误 ({sw.ElapsedMilliseconds}ms): {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"{method} 错误: {ex.Message}", "ERROR");
                if (respBox != null) respBox.Text = $"错误: {ex}";
                if (headerBox != null) headerBox.Text = "";
                AppendDetailedLog($"{method} Error", url, body, null, null, ex);
            }
        }

        /// <summary>
        /// 使用 SOAP/HTTP POST 调用 WebService 并显示结果（带日志记录）。
        /// 这是一个异步方法，UI 按钮事件应直接 await 或无返回地调用它。
        /// </summary>
        /// <param name="url">服务 URL</param>
        /// <param name="soapAction">SOAPAction 头（可选）</param>
        /// <param name="xmlBody">SOAP XML 请求体</param>
        private async Task DoWebServiceTest(string url, string soapAction, string xmlBody)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                MessageBox.Show("WebService：URL 为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                AppendLog($"WebService POST -> {url}");
                using var req = new HttpRequestMessage(HttpMethod.Post, url);
                req.Content = new StringContent(xmlBody ?? string.Empty, Encoding.UTF8, "text/xml");
                if (!string.IsNullOrWhiteSpace(soapAction)) req.Headers.Add("SOAPAction", soapAction);

                AppendDetailedLog("WebService Request", url + (string.IsNullOrEmpty(soapAction) ? "" : " SOAPAction=" + soapAction), xmlBody);

                using var resp = await _httpClient.SendAsync(req);
                var body = await resp.Content.ReadAsStringAsync();
                string result = $"状态: {(int)resp.StatusCode} {resp.ReasonPhrase}\n响应长度: {body?.Length ?? 0}";

                MessageBox.Show(result, "WebService 调用结果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                AppendLog($"状态: {(int)resp.StatusCode} {resp.ReasonPhrase}");
                AppendLog($"响应长度: {body?.Length ?? 0}");
                if (_wsResponseBox != null) _wsResponseBox.Text = body ?? string.Empty;

                AppendDetailedLog("WebService Response", url, null, $"{(int)resp.StatusCode} {resp.ReasonPhrase}", body);
            }
            catch (Exception ex)
            {
                string errorMsg = $"WebService 调用错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AppendLog($"WebService 调用错误: {ex.Message}", "ERROR");
                if (_wsResponseBox != null) _wsResponseBox.Text = $"错误: {ex}";
                AppendDetailedLog("WebService Error", url, xmlBody, null, null, ex);
            }
        }

        #region Protocol implementations with MessageBox results

        /// <summary>
        /// 将一条简短日志追加到 UI（_logBox）和内存 + 当日日志文件。
        /// 该方法是线程安全的，会在非 UI 线程时使用 BeginInvoke 更新 UI。
        /// </summary>
        /// <param name="text">日志文本</param>
        /// <param name="level">日志等级（默认 INFO）</param>
        private void AppendLog(string text, string level = "INFO")
        {
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {text}";
            lock (_logLock)
            {
                _logLines.Add(line);
            }

            WriteToDailyLog(line);

            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => { _logBox?.AppendText(line + Environment.NewLine); }));
            }
            else
            {
                _logBox?.AppendText(line + Environment.NewLine);
            }

            WriteToDailyLog(line);
        }

        /// <summary>
        /// 执行 HTTP GET 请求并显示响应信息，包含日志记录和错误处理。
        /// </summary>
        /// <param name="url">目标 URL</param>
        private async Task DoHttpGet(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                MessageBox.Show("HTTP GET：URL 为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                AppendLog($"HTTP GET -> {url}");
                AppendDetailedLog("HTTP GET Request", url);

                using var resp = await _httpClient.GetAsync(url);
                var body = await resp.Content.ReadAsStringAsync();
                string result = $"状态: {(int)resp.StatusCode} {resp.ReasonPhrase}\n响应长度: {body?.Length ?? 0}";

                MessageBox.Show(result, "HTTP GET 结果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                AppendLog($"状态: {(int)resp.StatusCode} {resp.ReasonPhrase}");
                AppendLog($"响应正文长度: {body?.Length ?? 0}");
                if (_httpResponseBox != null) _httpResponseBox.Text = body ?? "";

                AppendDetailedLog("HTTP GET Response", url, null, $"{(int)resp.StatusCode} {resp.ReasonPhrase}", body);
            }
            catch (Exception ex)
            {
                string errorMsg = $"HTTP GET 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AppendLog($"HTTP GET 错误: {ex.Message}", "ERROR");
                if (_httpResponseBox != null) _httpResponseBox.Text = $"错误: {ex}";
                AppendDetailedLog("HTTP GET Error", url, null, null, null, ex);
            }
        }

        /// <summary>
        /// 执行 HTTP POST（application/json）请求并显示响应，带日志记录。
        /// </summary>
        /// <param name="url">目标 URL</param>
        /// <param name="jsonBody">JSON 请求体</param>
        private async Task DoHttpPost(string url, string jsonBody)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                MessageBox.Show("HTTP POST：URL 为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                AppendLog($"HTTP POST -> {url}");
                var content = new StringContent(jsonBody ?? "", Encoding.UTF8, "application/json");

                AppendDetailedLog("HTTP POST Request", url, jsonBody);

                using var resp = await _httpClient.PostAsync(url, content);
                var body = await resp.Content.ReadAsStringAsync();
                string result = $"状态: {(int)resp.StatusCode} {resp.ReasonPhrase}\n响应长度: {body?.Length ?? 0}";

                MessageBox.Show(result, "HTTP POST 结果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                AppendLog($"状态: {(int)resp.StatusCode} {resp.ReasonPhrase}");
                AppendLog($"响应正文长度: {body?.Length ?? 0}");
                if (_httpResponseBox != null) _httpResponseBox.Text = body ?? "";

                AppendDetailedLog("HTTP POST Response", url, jsonBody, $"{(int)resp.StatusCode} {resp.ReasonPhrase}", body);
            }
            catch (Exception ex)
            {
                string errorMsg = $"HTTP POST 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AppendLog($"HTTP POST 错误: {ex.Message}", "ERROR");
                if (_httpResponseBox != null) _httpResponseBox.Text = $"错误: {ex}";
                AppendDetailedLog("HTTP POST Error", url, jsonBody, null, null, ex);
            }
        }

        /// <summary>
        /// 使用 ClientWebSocket 建立连接、发送字符串并接收回显（演示 WebSocket 使用）。
        /// 注：使用较短的超时时间并做基本错误处理。
        /// </summary>
        /// <param name="uri">WebSocket URI（ws:// 或 wss://）</param>
        private async Task DoWebSocketEcho(string uri)
        {
            if (!Uri.TryCreate(uri, UriKind.Absolute, out var u))
            {
                MessageBox.Show("WebSocket：无效 URI", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                AppendLog($"WebSocket -> {uri}");
                using var ws = new ClientWebSocket();
                var cts = new CancellationTokenSource(15000);
                await ws.ConnectAsync(u, cts.Token);

                MessageBox.Show("WebSocket：已连接", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                AppendLog("WebSocket：已连接");

                var sendBuf = Encoding.UTF8.GetBytes("hello from protocol debugger");
                await ws.SendAsync(sendBuf, WebSocketMessageType.Text, true, cts.Token);

                var recvBuf = new byte[8192];
                var res = await ws.ReceiveAsync(new ArraySegment<byte>(recvBuf), cts.Token);
                var msg = Encoding.UTF8.GetString(recvBuf, 0, res.Count);

                MessageBox.Show($"接收：{FirstChars(msg, 2000)}", "WebSocket 接收", MessageBoxButtons.OK, MessageBoxIcon.Information);
                AppendLog($"接收：{FirstChars(msg, 2000)}");

                await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "bye", CancellationToken.None);
                AppendLog("WebSocket：已关闭");
            }
            catch (Exception ex)
            {
                string errorMsg = $"WebSocket 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AppendLog($"WebSocket 错误: {ex.Message}", "ERROR");
                AppendDetailedLog("WebSocket Error", uri, null, null, null, ex);
            }
        }

        /// <summary>
        /// 使用 FtpWebRequest 列出 FTP 根目录（不支持复杂认证或 FTPS/TLS 的高级配置）。
        /// </summary>
        /// <param name="host">FTP 主机</param>
        /// <param name="portText">端口文本（未使用时忽略）</param>
        /// <param name="user">用户名</param>
        /// <param name="pass">密码</param>
        private async Task DoFtpList(string host, string portText, string user, string pass)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("FTP：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                var uri = new UriBuilder { Scheme = "ftp", Host = host, Path = "/" }.Uri;

                AppendLog($"FTP：列出 {uri}，用户 '{user}'");

                var request = (FtpWebRequest)WebRequest.Create(uri);
                request.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                if (!string.IsNullOrWhiteSpace(user)) request.Credentials = new NetworkCredential(user, pass);
                request.EnableSsl = false;

                using var resp = (FtpWebResponse)await request.GetResponseAsync();
                using var stream = resp.GetResponseStream();
                using var reader = new StreamReader(stream);
                var all = await reader.ReadToEndAsync();

                string result = $"FTP 状态: {resp.StatusDescription}\n{FirstChars(all, 8000)}";
                MessageBox.Show(result, "FTP 列表结果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                AppendLog($"FTP 状态: {resp.StatusDescription}");
                AppendLog(FirstChars(all, 8000));

                AppendDetailedLog("FTP List", uri.ToString(), null, resp.StatusDescription, all);
            }
            catch (Exception ex)
            {
                string errorMsg = $"FTP 列表错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AppendLog($"FTP 列表错误: {ex.Message}", "ERROR");
                AppendDetailedLog("FTP Error", host + ":" + portText, null, null, null, ex);
            }
        }

        /// <summary>
        /// 使用 Renci.SshNet 实现的 SFTP 列表功能（支持密码或私钥认证）。
        /// 注意：私钥可能需要密码保护，私钥加载失败会提示用户。
        /// </summary>
        /// <param name="host">SFTP 主机</param>
        /// <param name="portText">端口文本（解析为 int）</param>
        /// <param name="user">用户名</param>
        /// <param name="pass">密码</param>
        /// <param name="privateKeyPath">私钥路径（可选）</param>
        private async Task DoSftpList_Strong(string host, string portText, string user, string pass, string privateKeyPath)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("SFTP：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int port = 22;
            if (!string.IsNullOrWhiteSpace(portText) && int.TryParse(portText, out var p)) port = p;

            AppendLog($"SFTP -> {host}:{port}, user={user}, key={(string.IsNullOrEmpty(privateKeyPath) ? "<none>" : privateKeyPath)}");

            try
            {
                ConnectionInfo connInfo;

                if (!string.IsNullOrWhiteSpace(privateKeyPath) && File.Exists(privateKeyPath))
                {
                    var keyFiles = new List<PrivateKeyFile>();
                    try { keyFiles.Add(new PrivateKeyFile(privateKeyPath)); }
                    catch (Exception)
                    {
                        MessageBox.Show("SFTP：私钥加载失败（可能需要带密码的私钥）。若私钥受保护，请同时填写密码字段。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    var methods = new List<AuthenticationMethod>();
                    if (!string.IsNullOrEmpty(pass)) methods.Add(new PasswordAuthenticationMethod(user, pass));
                    if (keyFiles.Count > 0) methods.Add(new PrivateKeyAuthenticationMethod(user, keyFiles.ToArray()));

                    connInfo = new ConnectionInfo(host, port, user, methods.ToArray());
                }
                else
                {
                    connInfo = new PasswordConnectionInfo(host, port, user, pass);
                }

                using var sftp = new SftpClient(connInfo);
                sftp.Connect();
                if (!sftp.IsConnected)
                {
                    MessageBox.Show("SFTP：连接失败", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                AppendLog("SFTP：已连接，列出根目录...");
                var files = sftp.ListDirectory("/");
                var sb = new StringBuilder();
                foreach (var f in files) sb.AppendLine($"{f.Name}\t{(f.IsDirectory ? "<DIR>" : f.Length.ToString())}\t{f.LastWriteTime}");

                string result = sb.ToString();
                MessageBox.Show(FirstChars(result, 8000), "SFTP 列表结果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                AppendLog(FirstChars(sb.ToString(), 8000));
                sftp.Disconnect();
                AppendLog("SFTP：已断开");

                AppendDetailedLog("SFTP List", host + ":" + port, null, "Listed", sb.ToString());
            }
            catch (Exception ex)
            {
                string errorMsg = $"SFTP 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AppendLog($"SFTP 错误: {ex.Message}", "ERROR");
                AppendDetailedLog("SFTP Error", host + ":" + port, null, null, null, ex);
            }

            // 该方法为异步签名，但 SftpClient 的连接/列出是同步 API，方法最后通过 Task.Yield 保持异步契约
            await Task.Yield();
        }

        /// <summary>
        /// FTP 上传文件。
        /// </summary>
        private async Task DoFtpUpload(string host, string portText, string user, string pass, string localPath, string remotePath)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("FTP 上传：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
            {
                MessageBox.Show("FTP 上传：本地文件不存在", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var port = string.IsNullOrWhiteSpace(portText) ? "" : ":" + portText;
                var uri = $"ftp://{host}{port}{remotePath}";
                AppendLog($"FTP 上传: {localPath} -> {uri}");

                var request = (FtpWebRequest)WebRequest.Create(uri);
                request.Method = WebRequestMethods.Ftp.UploadFile;
                if (!string.IsNullOrWhiteSpace(user)) request.Credentials = new NetworkCredential(user, pass);
                request.UseBinary = true;
                request.EnableSsl = false;

                var fileBytes = await File.ReadAllBytesAsync(localPath);
                using var reqStream = await request.GetRequestStreamAsync();
                await reqStream.WriteAsync(fileBytes, 0, fileBytes.Length);

                using var resp = (FtpWebResponse)await request.GetResponseAsync();
                string result = $"FTP 上传成功: {fileBytes.Length} 字节\n状态: {resp.StatusDescription}";
                MessageBox.Show(result, "FTP 上传", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (_ftpResponseBox != null) _ftpResponseBox.Text = result;
                AppendLog($"FTP 上传成功: {fileBytes.Length} 字节");
            }
            catch (Exception ex)
            {
                string errorMsg = $"FTP 上传错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"FTP 上传错误: {ex.Message}", "ERROR");
                if (_ftpResponseBox != null) _ftpResponseBox.Text = errorMsg;
            }

            await Task.Yield();
        }

        /// <summary>
        /// FTP 下载文件。
        /// </summary>
        private async Task DoFtpDownload(string host, string portText, string user, string pass, string remotePath, string localPath)
        {
            if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(remotePath))
            {
                MessageBox.Show("FTP 下载：主机或远程路径为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var port = string.IsNullOrWhiteSpace(portText) ? "" : ":" + portText;
                var uri = $"ftp://{host}{port}{remotePath}";
                AppendLog($"FTP 下载: {uri} -> {localPath}");

                var request = (FtpWebRequest)WebRequest.Create(uri);
                request.Method = WebRequestMethods.Ftp.DownloadFile;
                if (!string.IsNullOrWhiteSpace(user)) request.Credentials = new NetworkCredential(user, pass);
                request.UseBinary = true;
                request.EnableSsl = false;

                using var resp = (FtpWebResponse)await request.GetResponseAsync();
                using var stream = resp.GetResponseStream();
                using var ms = new MemoryStream();
                await stream.CopyToAsync(ms);
                var fileBytes = ms.ToArray();

                await File.WriteAllBytesAsync(localPath, fileBytes);
                string result = $"FTP 下载成功: {fileBytes.Length} 字节 -> {localPath}\n状态: {resp.StatusDescription}";
                MessageBox.Show(result, "FTP 下载", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (_ftpResponseBox != null) _ftpResponseBox.Text = result;
                AppendLog($"FTP 下载成功: {fileBytes.Length} 字节");
            }
            catch (Exception ex)
            {
                string errorMsg = $"FTP 下载错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"FTP 下载错误: {ex.Message}", "ERROR");
                if (_ftpResponseBox != null) _ftpResponseBox.Text = errorMsg;
            }

            await Task.Yield();
        }

        /// <summary>
        /// FTP 删除文件。
        /// </summary>
        private async Task DoFtpDelete(string host, string portText, string user, string pass, string remotePath)
        {
            if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(remotePath))
            {
                MessageBox.Show("FTP 删除：主机或远程路径为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var port = string.IsNullOrWhiteSpace(portText) ? "" : ":" + portText;
                var uri = $"ftp://{host}{port}{remotePath}";
                AppendLog($"FTP 删除: {uri}");

                var request = (FtpWebRequest)WebRequest.Create(uri);
                request.Method = WebRequestMethods.Ftp.DeleteFile;
                if (!string.IsNullOrWhiteSpace(user)) request.Credentials = new NetworkCredential(user, pass);
                request.EnableSsl = false;

                using var resp = (FtpWebResponse)await request.GetResponseAsync();
                string result = $"FTP 删除成功: {remotePath}\n状态: {resp.StatusDescription}";
                MessageBox.Show(result, "FTP 删除", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (_ftpResponseBox != null) _ftpResponseBox.Text = result;
                AppendLog($"FTP 删除成功: {remotePath}");
            }
            catch (Exception ex)
            {
                string errorMsg = $"FTP 删除错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"FTP 删除错误: {ex.Message}", "ERROR");
                if (_ftpResponseBox != null) _ftpResponseBox.Text = errorMsg;
            }

            await Task.Yield();
        }

        /// <summary>
        /// FTP 创建目录。
        /// </summary>
        private async Task DoFtpMkdir(string host, string portText, string user, string pass, string remotePath)
        {
            if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(remotePath))
            {
                MessageBox.Show("FTP 创建目录：主机或路径为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var port = string.IsNullOrWhiteSpace(portText) ? "" : ":" + portText;
                var uri = $"ftp://{host}{port}{remotePath}";
                AppendLog($"FTP 创建目录: {uri}");

                var request = (FtpWebRequest)WebRequest.Create(uri);
                request.Method = WebRequestMethods.Ftp.MakeDirectory;
                if (!string.IsNullOrWhiteSpace(user)) request.Credentials = new NetworkCredential(user, pass);
                request.EnableSsl = false;

                using var resp = (FtpWebResponse)await request.GetResponseAsync();
                string result = $"FTP 创建目录成功: {remotePath}\n状态: {resp.StatusDescription}";
                MessageBox.Show(result, "FTP 创建目录", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (_ftpResponseBox != null) _ftpResponseBox.Text = result;
                AppendLog($"FTP 创建目录成功: {remotePath}");
            }
            catch (Exception ex)
            {
                string errorMsg = $"FTP 创建目录错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"FTP 创建目录错误: {ex.Message}", "ERROR");
                if (_ftpResponseBox != null) _ftpResponseBox.Text = errorMsg;
            }

            await Task.Yield();
        }

        /// <summary>
        /// SSH 远程执行命令，显示 stdout/stderr 和退出码。
        /// </summary>
        private async Task DoSshCommand(string host, string portText, string user, string pass, string command, TextBox respBox)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("SSH：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(command))
            {
                MessageBox.Show("SSH：命令为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int port = 22;
            if (!string.IsNullOrWhiteSpace(portText) && int.TryParse(portText, out var p)) port = p;

            AppendLog($"SSH -> {host}:{port} user={user} cmd={FirstChars(command, 100)}");
            try
            {
                using var client = new SshClient(host, port, user, pass);
                client.Connect();
                if (!client.IsConnected)
                {
                    MessageBox.Show("SSH 连接失败", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    AppendLog("SSH 连接失败", "ERROR");
                    return;
                }

                AppendLog("SSH：已连接，执行命令...");
                var cmd = client.RunCommand(command);
                var sb = new StringBuilder();
                sb.AppendLine($"=== SSH 命令结果 ===");
                sb.AppendLine($"命令: {command}");
                sb.AppendLine($"退出码: {cmd.ExitStatus}");
                sb.AppendLine();

                if (!string.IsNullOrEmpty(cmd.Result))
                {
                    sb.AppendLine("--- STDOUT ---");
                    sb.AppendLine(cmd.Result);
                }
                if (!string.IsNullOrEmpty(cmd.Error))
                {
                    sb.AppendLine("--- STDERR ---");
                    sb.AppendLine(cmd.Error);
                }

                string result = sb.ToString();
                if (respBox != null) respBox.Text = result;
                AppendLog($"SSH 命令完成，退出码={cmd.ExitStatus}");
                AppendDetailedLog("SSH Command", host + ":" + port, command, $"ExitCode={cmd.ExitStatus}", result);

                client.Disconnect();
                AppendLog("SSH：已断开");
            }
            catch (Exception ex)
            {
                string errorMsg = $"SSH 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"SSH 错误: {ex.Message}", "ERROR");
                if (respBox != null) respBox.Text = errorMsg;
                AppendDetailedLog("SSH Error", host + ":" + port, command, null, null, ex);
            }

            await Task.Yield();
        }

        /// <summary>
        /// 使用 System.Net.Mail.SmtpClient 发送邮件（最简单的示例，适合测试）。
        /// 注意：SmtpClient 在某些场景下已被标记为过时，但对于简单发送仍可用；生产建议使用 MailKit 等库。
        /// </summary>
        private async Task DoSendSmtp(string host, string portText, string user, string pass, string from, string to, string subject, string body, bool EnableSsl)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("SMTP：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                var port = 25;
                if (!string.IsNullOrWhiteSpace(portText) && int.TryParse(portText, out var p)) port = p;

                AppendLog($"SMTP -> {host}:{port} 发件人={from} 收件人={to}");

                var msg = new MailMessage(from, to, subject, body);
                AppendDetailedLog("SMTP Send", host + ":" + port, $"From:{from}\nTo:{to}\nSubject:{subject}\nBody:\n{body}");

                using var client = new SmtpClient(host, port) { EnableSsl = EnableSsl };
                if (!string.IsNullOrWhiteSpace(user)) client.Credentials = new NetworkCredential(user, pass);

                await client.SendMailAsync(msg);

                MessageBox.Show("SMTP：发送成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                AppendLog("SMTP：发送成功（SmtpClient）");
                AppendDetailedLog("SMTP Sent", host + ":" + port, null, "Sent", null);
            }
            catch (Exception ex)
            {
                string errorMsg = $"SMTP 发送错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AppendLog($"SMTP 发送错误: {ex.Message}", "ERROR");
                AppendDetailedLog("SMTP Error", host + ":" + portText, null, null, null, ex);
            }
        }

        /// <summary>
        /// 发送 REST POST（application/json），并把响应写入对应的响应框。
        /// </summary>
        private async Task DoRestPost(string url, string json)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                MessageBox.Show("REST POST：URL 为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                AppendLog($"REST POST -> {url}");
                var content = new StringContent(json ?? "", Encoding.UTF8, "application/json");

                AppendDetailedLog("REST POST Request", url, json);

                using var resp = await _httpClient.PostAsync(url, content);
                var body = await resp.Content.ReadAsStringAsync();
                string result = $"状态: {(int)resp.StatusCode} {resp.ReasonPhrase}\n响应长度: {body?.Length ?? 0}";

                MessageBox.Show(result, "REST POST 结果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                AppendLog($"状态: {(int)resp.StatusCode} {resp.ReasonPhrase}");
                AppendLog($"响应正文长度: {body?.Length ?? 0}");
                if (_restResponseBox != null) _restResponseBox.Text = body ?? "";

                AppendDetailedLog("REST POST Response", url, json, $"{(int)resp.StatusCode} {resp.ReasonPhrase}", body);
            }
            catch (Exception ex)
            {
                string errorMsg = $"REST POST 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AppendLog($"REST POST 错误: {ex.Message}", "ERROR");
                if (_restResponseBox != null) _restResponseBox.Text = $"错误: {ex}";
                AppendDetailedLog("REST POST Error", url, json, null, null, ex);
            }
        }

        /// <summary>
        /// 使用 MailKit 的 Pop3Client 获取邮件摘要（演示如何使用 MailKit）。
        /// </summary>
        private async Task DoPop3MailKit(string host, string portText, string user, string pass, bool useSsl)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("POP3：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(portText) || !int.TryParse(portText, out var port)) port = useSsl ? 995 : 110;

            AppendLog($"POP3 -> {host}:{port} SSL={useSsl} user={user}");
            try
            {
                using var client = new Pop3Client();
                var socketOptions = useSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTlsWhenAvailable;
                await client.ConnectAsync(host, port, socketOptions);
                if (!string.IsNullOrEmpty(user)) await client.AuthenticateAsync(user, pass);

                var count = client.Count;
                string result = $"POP3：邮件总数 {count}";
                MessageBox.Show(result, "POP3 结果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                AppendLog($"POP3：邮件总数 {count}");
                var max = Math.Min(5, count);
                for (int i = 0; i < max; i++)
                {
                    var msg = await client.GetMessageAsync(i);
                    AppendLog($"[{i + 1}] {msg.Date}: {msg.Subject} From: {string.Join(", ", msg.From.Select(x => x.ToString()))}");
                }

                await client.DisconnectAsync(true);
                AppendLog("POP3：已断开");

                AppendDetailedLog("POP3 Summary", host + ":" + port, null, $"Count={count}", null);
            }
            catch (Exception ex)
            {
                string errorMsg = $"POP3 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AppendLog($"POP3 错误: {ex.Message}", "ERROR");
                AppendDetailedLog("POP3 Error", host + ":" + portText, null, null, null, ex);
            }
        }

        /// <summary>
        /// 使用 MailKit 的 ImapClient 获取收件箱摘要。
        /// </summary>
        private async Task DoImapMailKit(string host, string portText, string user, string pass, bool useSsl)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("IMAP：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(portText) || !int.TryParse(portText, out var port)) port = useSsl ? 993 : 143;

            AppendLog($"IMAP -> {host}:{port} SSL={useSsl} user={user}");
            try
            {
                using var client = new ImapClient();
                var socketOptions = useSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTlsWhenAvailable;
                await client.ConnectAsync(host, port, socketOptions);
                if (!string.IsNullOrEmpty(user)) await client.AuthenticateAsync(user, pass);

                var inbox = client.Inbox;
                await inbox.OpenAsync(FolderAccess.ReadOnly);
                string result = $"IMAP：收件箱共 {inbox.Count} 封（最近 {Math.Min(10, inbox.Count)} 封列出）";
                MessageBox.Show(result, "IMAP 结果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                AppendLog($"IMAP：收件箱共 {inbox.Count} 封（最近 {Math.Min(10, inbox.Count)} 封列出）");

                var start = Math.Max(0, inbox.Count - 10);
                for (int i = inbox.Count - 1; i >= start; i--)
                {
                    var message = await inbox.GetMessageAsync(i);
                    AppendLog($"[{i + 1}] {message.Date}: {message.Subject} From: {string.Join(", ", message.From.Select(x => x.ToString()))}");
                }

                await client.DisconnectAsync(true);
                AppendLog("IMAP：已断开");

                AppendDetailedLog("IMAP Summary", host + ":" + port, null, $"Count={inbox.Count}", null);
            }
            catch (Exception ex)
            {
                string errorMsg = $"IMAP 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AppendLog($"IMAP 错误: {ex.Message}", "ERROR");
                AppendDetailedLog("IMAP Error", host + ":" + portText, null, null, null, ex);
            }
        }

        /// <summary>
        /// 测试目标主机端口连通性（TCP 连接），返回结果描述文本并弹窗显示。
        /// 超时时间为 5 秒。
        /// </summary>
        /// <returns>用于显示的结果字符串或错误文本</returns>
        private async Task<string> TestIpPort(string hostOrIp, string portText)
        {
            if (string.IsNullOrWhiteSpace(hostOrIp) || string.IsNullOrWhiteSpace(portText) || !int.TryParse(portText, out var port))
            {
                string errorMsg = "IP:Port 测试：主机或端口无效";
                MessageBox.Show(errorMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                AppendLog("IP:Port 测试：主机或端口无效", "WARN");
                return "主机或端口无效";
            }

            AppendLog($"测试连接 -> {hostOrIp}:{port}");
            var sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                using var tcp = new TcpClient();
                var connectTask = tcp.ConnectAsync(hostOrIp, port);
                if (await Task.WhenAny(connectTask, Task.Delay(5000)) == connectTask)
                {
                    sw.Stop();
                    string successMsg = $"IP:Port 可达（TCP 连接成功）\n目标：{hostOrIp}:{port}\n耗时：{sw.ElapsedMilliseconds} ms";
                    MessageBox.Show(successMsg, "IP:Port 检测结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    AppendLog($"IP:Port 可达（TCP 连接成功），耗时 {sw.ElapsedMilliseconds} ms");
                    return successMsg;
                }
                else
                {
                    sw.Stop();
                    string timeoutMsg = $"IP:Port 不可达（连接超时）\n目标：{hostOrIp}:{port}\n耗时：{sw.ElapsedMilliseconds} ms";
                    MessageBox.Show(timeoutMsg, "IP:Port 检测结果", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    AppendLog($"IP:Port 不可达（超时），耗时 {sw.ElapsedMilliseconds} ms", "WARN");
                    return timeoutMsg;
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"IP:Port 测试错误\n目标：{hostOrIp}:{port}\n错误信息：{ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"IP:Port 测试错误: {ex.Message}", "ERROR");
                AppendDetailedLog("IP:Port Error", hostOrIp + ":" + portText, null, null, null, ex);
                return errorMsg;
            }
        }

        /// <summary>
        /// 构建 SQL Server 连接字符串（尽可能容错）。
        /// 生产环境请使用官方连接字符串构建器类（如 SqlConnectionStringBuilder）。
        /// </summary>
        private string BuildSqlServerConnectionString(string host, string port, string database, string user, string pass)
        {
            if (string.IsNullOrWhiteSpace(host)) return "";
            var builder = new StringBuilder();
            builder.Append($"Server={host}");
            if (!string.IsNullOrWhiteSpace(port)) builder.Append($",{port}");
            if (!string.IsNullOrWhiteSpace(database)) builder.Append($";Database={database}");
            if (!string.IsNullOrWhiteSpace(user))
                builder.Append($";User Id={user};Password={pass};");
            else
                builder.Append(";Integrated Security=true;");
            builder.Append("TrustServerCertificate=true;");
            return builder.ToString();
        }

        /// <summary>
        /// 构建 Oracle 连接字符串（简化版本）。
        /// </summary>
        private string BuildOracleConnectionString(string host, string port, string serviceName, string user, string pass)
        {
            if (string.IsNullOrWhiteSpace(host)) return "";
            var sb = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(user))
            {
                sb.Append($"User Id={user};Password={pass};");
            }
            var hostPort = string.IsNullOrWhiteSpace(port) ? "1521" : port;
            sb.Append($"Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={hostPort}))(CONNECT_DATA=(SERVICE_NAME={serviceName})));");
            sb.Append("Persist Security Info=True;");
            return sb.ToString();
        }

        /// <summary>
        /// 构建 MySQL 连接字符串（简化版本）。
        /// </summary>
        private string BuildMySqlConnectionString(string host, string port, string database, string user, string pass)
        {
            if (string.IsNullOrWhiteSpace(host)) return "";
            var sb = new StringBuilder();
            sb.Append($"Server={host};");
            if (!string.IsNullOrWhiteSpace(port)) sb.Append($"Port={port};");
            if (!string.IsNullOrWhiteSpace(database)) sb.Append($"Database={database};");
            sb.Append($"User Id={user};Password={pass};");
            sb.Append("SslMode=Preferred;");
            return sb.ToString();
        }

        /// <summary>
        /// 构建 PostgreSQL 连接字符串（简化版本）。
        /// </summary>
        private string BuildNpgsqlConnectionString(string host, string port, string database, string user, string pass)
        {
            if (string.IsNullOrWhiteSpace(host)) return "";
            var sb = new StringBuilder();
            sb.Append($"Host={host};");
            if (!string.IsNullOrWhiteSpace(port)) sb.Append($"Port={port};");
            if (!string.IsNullOrWhiteSpace(database)) sb.Append($"Database={database};");
            sb.Append($"Username={user};Password={pass};");
            return sb.ToString();
        }

        /// <summary>
        /// 测试数据库连接（支持通过反射加载不同的 ADO.NET 驱动并尝试 Open / OpenAsync）。
        /// 此函数会在 UI 上弹窗并将结果写入日志与结果表格。
        /// </summary>
        /// <param name="providerKey">驱动标识（如 sqlserver/oracle/mysql）</param>
        /// <param name="connectionString">完整连接字符串</param>
        private async Task TestDbConnection(string providerKey, string connectionString)
        {
            var timestamp = DateTime.Now;
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                string errorMsg = $"[{providerKey}] 连接字符串为空，跳过测试。";
                MessageBox.Show(errorMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                AppendLog($"[{providerKey}] 连接字符串为空，跳过测试。", "WARN");
                AddDbResultRow(timestamp, providerKey, "-", "Skipped", 0, "连接字符串为空");
                return;
            }

            AppendLog($"[{providerKey}] 测试连接，连接字符串长度 {connectionString.Length}");
            var sw = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                Type connType = GetConnectionTypeForProvider(providerKey);
                if (connType == null)
                {
                    var advice = providerKey switch
                    {
                        "sqlserver" => "Install-Package Microsoft.Data.SqlClient",
                        "oracle" => "Install-Package Oracle.ManagedDataAccess.Core",
                        "mysql" => "Install-Package MySql.Data 或 Install-Package MySqlConnector",
                        "npgsql" => "Install-Package Npgsql",
                        _ => "请安装相应驱动"
                    };
                    string errorMsg = $"[{providerKey}] 未找到对应的数据库驱动。请通过 Package Manager Console 安装对应 NuGet 包： {advice}";
                    MessageBox.Show(errorMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    AppendLog($"[{providerKey}] 未找到对应的数据库驱动。请通过 __Package Manager Console__ 安装对应 NuGet 包： {advice}", "WARN");
                    AddDbResultRow(timestamp, providerKey, MaskConnectionString(connectionString), "DriverMissing", sw.ElapsedMilliseconds, advice);
                    return;
                }

                object connInstance = null;

                var ctor = connType.GetConstructor(new[] { typeof(string) });
                if (ctor != null)
                {
                    connInstance = ctor.Invoke(new object[] { connectionString });
                }
                else
                {
                    connInstance = Activator.CreateInstance(connType);
                    var prop = connType.GetProperty("ConnectionString");
                    if (prop != null) prop.SetValue(connInstance, connectionString);
                }

                // 尝试 OpenAsync -> Open -> 报错处理（兼容不同驱动实现）
                var openAsyncMethod = connType.GetMethod("OpenAsync", Type.EmptyTypes) ?? connType.GetMethod("OpenAsync", new[] { typeof(CancellationToken) });
                if (openAsyncMethod != null)
                {
                    var taskObj = openAsyncMethod.GetParameters().Length == 0
                        ? (Task)openAsyncMethod.Invoke(connInstance, null)
                        : (Task)openAsyncMethod.Invoke(connInstance, new object[] { CancellationToken.None });
                    await taskObj;
                }
                else
                {
                    var openMethod = connType.GetMethod("Open", Type.EmptyTypes);
                    if (openMethod != null)
                    {
                        await Task.Run(() => openMethod.Invoke(connInstance, null));
                    }
                    else
                    {
                        string errorMsg = $"[{providerKey}] 未找到 Open/OpenAsync 方法，无法测试连接。";
                        MessageBox.Show(errorMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        AppendLog($"[{providerKey}] 未找到 Open/OpenAsync 方法，无法测试连接。", "WARN");
                        AddDbResultRow(timestamp, providerKey, MaskConnectionString(connectionString), "NoOpenMethod", sw.ElapsedMilliseconds, "未找到 Open/OpenAsync");
                        return;
                    }
                }

                sw.Stop();
                string successMsg = $"[{providerKey}] 连接成功。耗时 {sw.ElapsedMilliseconds} ms";
                MessageBox.Show(successMsg, "数据库连接测试", MessageBoxButtons.OK, MessageBoxIcon.Information);
                AppendLog($"[{providerKey}] 连接成功。耗时 {sw.ElapsedMilliseconds} ms");
                AddDbResultRow(timestamp, providerKey, MaskConnectionString(connectionString), "Success", sw.ElapsedMilliseconds, "连接成功");

                var closeMethod = connType.GetMethod("Close") ?? connType.GetMethod("Dispose");
                closeMethod?.Invoke(connInstance, null);
                (connInstance as IDisposable)?.Dispose();
            }
            catch (TargetInvocationException tie)
            {
                sw.Stop();
                var msg = tie.InnerException?.Message ?? tie.Message;
                string errorMsg = $"[{providerKey}] 驱动抛出异常: {msg}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"[{providerKey}] 驱动抛出异常: {msg}", "ERROR");
                AddDbResultRow(timestamp, providerKey, MaskConnectionString(connectionString), "Error", sw.ElapsedMilliseconds, msg);
            }
            catch (Exception ex)
            {
                sw.Stop();
                string errorMsg = $"[{providerKey}] 连接失败: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"[{providerKey}] 连接失败: {ex.Message}", "ERROR");
                AddDbResultRow(timestamp, providerKey, MaskConnectionString(connectionString), "Fail", sw.ElapsedMilliseconds, ex.Message);
            }
        }

        /// <summary>
        /// 根据 providerKey 返回对应的连接类型（反射查找），支持多种常见驱动名称。
        /// 返回 null 表示未找到对应驱动。
        /// </summary>
        private Type GetConnectionTypeForProvider(string providerKey)
        {
            if (providerKey == "sqlserver")
            {
                var t = Type.GetType("Microsoft.Data.SqlClient.SqlConnection, Microsoft.Data.SqlClient");
                if (t != null) return t;
                t = Type.GetType("System.Data.SqlClient.SqlConnection, System.Data");
                if (t != null) return t;
                return null;
            }

            if (providerKey == "oracle")
            {
                var candidates = new[]
                {
                    "Oracle.ManagedDataAccess.Client.OracleConnection, Oracle.ManagedDataAccess.Core",
                    "Oracle.ManagedDataAccess.Client.OracleConnection, Oracle.ManagedDataAccess"
                };
                foreach (var s in candidates)
                {
                    var t = Type.GetType(s);
                    if (t != null) return t;
                }
                return null;
            }

            if (providerKey == "mysql")
            {
                var t = Type.GetType("MySql.Data.MySqlClient.MySqlConnection, MySql.Data");
                if (t != null) return t;
                t = Type.GetType("MySqlConnector.MySqlConnection, MySqlConnector");
                if (t != null) return t;
                return null;
            }

            if (providerKey == "npgsql")
            {
                var t = Type.GetType("Npgsql.NpgsqlConnection, Npgsql");
                if (t != null) return t;
                return null;
            }

            return null;
        }

        /// <summary>
        /// 对连接字符串做简单掩码处理，隐藏 password 字段以避免日志泄露敏感信息。
        /// 如果字符串过长，会进行截断显示。
        /// </summary>
        private string MaskConnectionString(string cs)
        {
            if (string.IsNullOrEmpty(cs)) return "";
            var lower = cs.ToLowerInvariant();
            var idx = lower.IndexOf("password=");
            if (idx >= 0)
            {
                var sb = new StringBuilder(cs);
                var start = idx + "password=".Length;
                var end = cs.IndexOf(';', start);
                if (end < 0) end = cs.Length;
                for (int i = start; i < end; i++) sb[i] = '*';
                return sb.ToString();
            }
            return cs.Length > 120 ? cs.Substring(0, 120) + "..." : cs;
        }

        /// <summary>
        /// 将数据库检测结果插入结果表格（如果存在）。
        /// 该方法会在 UI 线程执行插入（若当前不在 UI 线程则使用 BeginInvoke）。
        /// </summary>
        private void AddDbResultRow(DateTime time, string dbType, string target, string result, long elapsedMs, string message)
        {
            if (_dbResultGrid == null) return;

            var row = new string[]
            {
                time.ToString("yyyy-MM-dd HH:mm:ss"),
                dbType,
                target?.Length > 100 ? target.Substring(0,100) + "..." : target,
                result,
                elapsedMs.ToString(),
                message
            };

            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => _dbResultGrid.Rows.Insert(0, row)));
                return;
            }

            _dbResultGrid.Rows.Insert(0, row);
        }

        /// <summary>
        /// 简单的 Modbus TCP 读取示例（手动构造 MBAP + PDU），仅用于测试和调试。
        /// 注意：此实现为最小示例，未实现完整的异常码/异常 PDU 解析。
        /// </summary>
        private async Task TestModbusTcp(string host, string portText, string unitIdText, string startText, string qtyText)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("Modbus 测试：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!int.TryParse(portText, out var port)) port = 502;
            if (!byte.TryParse(unitIdText, out var unitId)) unitId = 1;
            if (!ushort.TryParse(startText, out var start)) start = 0;
            if (!ushort.TryParse(qtyText, out var qty)) qty = 1;

            AppendLog($"Modbus TCP -> {host}:{port} Unit={unitId} Start={start} Qty={qty}");
            try
            {
                using var tcp = new TcpClient();
                var connectTask = tcp.ConnectAsync(host, port);
                if (await Task.WhenAny(connectTask, Task.Delay(4000)) != connectTask)
                {
                    string errorMsg = "Modbus: TCP 连接超时";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _deviceProtoResponseBox.Text = "Modbus: TCP 连接超时";
                    AppendLog("Modbus: TCP 连接超时", "WARN");
                    return;
                }

                using var stream = tcp.GetStream();
                var transId = (ushort)new Random().Next(1, ushort.MaxValue);
                var protocolId = (ushort)0;
                var pdu = new List<byte>();
                pdu.Add(0x03);
                pdu.Add((byte)(start >> 8));
                pdu.Add((byte)(start & 0xFF));
                pdu.Add((byte)(qty >> 8));
                pdu.Add((byte)(qty & 0xFF));
                var length = (ushort)(pdu.Count + 1);
                var adu = new List<byte>();
                adu.Add((byte)(transId >> 8));
                adu.Add((byte)(transId & 0xFF));
                adu.Add((byte)(protocolId >> 8));
                adu.Add((byte)(protocolId & 0xFF));
                adu.Add((byte)(length >> 8));
                adu.Add((byte)(length & 0xFF));
                adu.Add(unitId);
                adu.AddRange(pdu);

                await stream.WriteAsync(adu.ToArray(), 0, adu.Count);
                var header = new byte[7];
                var read = await stream.ReadAsync(header, 0, 7);
                if (read < 7)
                {
                    string errorMsg = "Modbus: 响应头读取失败";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _deviceProtoResponseBox.Text = "Modbus: 响应头读取失败";
                    AppendLog("Modbus: 响应头读取失败", "WARN");
                    return;
                }
                var respTrans = (ushort)(header[0] << 8 | header[1]);
                var respLen = (ushort)(header[4] << 8 | header[5]);
                var respUnit = header[6];
                var body = new byte[respLen - 1];
                var offset = 0;
                while (offset < body.Length)
                {
                    var r = await stream.ReadAsync(body, offset, body.Length - offset);
                    if (r <= 0) break;
                    offset += r;
                }

                var sb = new StringBuilder();
                sb.AppendLine($"TransId={respTrans}, Unit={respUnit}, BodyLen={body.Length}");
                sb.AppendLine("Body (hex): " + BitConverter.ToString(body));

                string result = sb.ToString();
                MessageBox.Show(result, "Modbus 响应", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _deviceProtoResponseBox.Text = result;
                AppendLog("Modbus: 收到响应 " + FirstChars(sb.ToString(), 800));
            }
            catch (Exception ex)
            {
                string errorMsg = $"Modbus 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _deviceProtoResponseBox.Text = $"Modbus 错误: {ex.Message}";
                AppendLog($"Modbus 错误: {ex.Message}", "ERROR");
            }
        }

        /// <summary>
        /// Modbus TCP 写操作（支持 FC05 写线圈、FC06 写单个寄存器、FC16 写多个寄存器）。
        /// </summary>
        /// <param name="host">目标主机</param>
        /// <param name="portText">端口（默认 502）</param>
        /// <param name="unitIdText">Unit ID</param>
        /// <param name="functionCode">功能码：5=写线圈, 6=写单寄存器, 16=写多寄存器</param>
        /// <param name="startText">起始地址</param>
        /// <param name="valuesText">写入值（多个用逗号分隔）</param>
        /// <param name="respBox">结果显示控件</param>
        private async Task TestModbusWrite(string host, string portText, string unitIdText, string functionCode, string startText, string valuesText, TextBox respBox)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("Modbus 写入：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!int.TryParse(portText, out var port)) port = 502;
            if (!byte.TryParse(unitIdText, out var unitId)) unitId = 1;
            if (!ushort.TryParse(startText, out var start)) start = 0;

            AppendLog($"Modbus Write -> {host}:{port} Unit={unitId} FC={functionCode} Start={start} Values={valuesText}");
            try
            {
                using var tcp = new TcpClient();
                var connectTask = tcp.ConnectAsync(host, port);
                if (await Task.WhenAny(connectTask, Task.Delay(4000)) != connectTask)
                {
                    string errorMsg = "Modbus 写入: TCP 连接超时";
                    if (respBox != null) respBox.Text = errorMsg;
                    AppendLog(errorMsg, "WARN");
                    return;
                }

                using var stream = tcp.GetStream();
                var transId = (ushort)Random.Shared.Next(1, ushort.MaxValue);

                List<byte> pdu;
                byte fc = functionCode switch
                {
                    "5" or "FC05" or "WriteCoil" => 0x05,
                    "6" or "FC06" or "WriteRegister" => 0x06,
                    "16" or "FC16" or "WriteMultiple" => 0x10,
                    _ => 0x06
                };

                var vals = (valuesText ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(v => v.Trim()).ToArray();

                if (fc == 0x05) // 写单个线圈
                {
                    ushort coilValue = (vals.Length > 0 && vals[0] == "0") ? (ushort)0x0000 : (ushort)0xFF00;
                    pdu = new List<byte> { fc, (byte)(start >> 8), (byte)(start & 0xFF), (byte)(coilValue >> 8), (byte)(coilValue & 0xFF) };
                }
                else if (fc == 0x06) // 写单个寄存器
                {
                    ushort val = vals.Length > 0 && ushort.TryParse(vals[0], out var v) ? v : (ushort)0;
                    pdu = new List<byte> { fc, (byte)(start >> 8), (byte)(start & 0xFF), (byte)(val >> 8), (byte)(val & 0xFF) };
                }
                else // FC16 写多个寄存器
                {
                    var registers = vals.Select(v => ushort.TryParse(v, out var r) ? r : (ushort)0).ToArray();
                    ushort qty = (ushort)registers.Length;
                    byte byteCount = (byte)(qty * 2);
                    pdu = new List<byte> { fc, (byte)(start >> 8), (byte)(start & 0xFF), (byte)(qty >> 8), (byte)(qty & 0xFF), byteCount };
                    foreach (var reg in registers) { pdu.Add((byte)(reg >> 8)); pdu.Add((byte)(reg & 0xFF)); }
                }

                // 构建 MBAP 头 + PDU
                ushort length = (ushort)(pdu.Count + 1);
                var adu = new List<byte>();
                adu.Add((byte)(transId >> 8)); adu.Add((byte)(transId & 0xFF));
                adu.Add(0); adu.Add(0); // Protocol ID
                adu.Add((byte)(length >> 8)); adu.Add((byte)(length & 0xFF));
                adu.Add(unitId);
                adu.AddRange(pdu);

                await stream.WriteAsync(adu.ToArray(), 0, adu.Count);

                // 读取响应头（7 字节）
                var header = new byte[7];
                var read = await stream.ReadAsync(header, 0, 7);
                if (read < 7)
                {
                    string errorMsg = "Modbus 写入: 响应头读取失败";
                    if (respBox != null) respBox.Text = errorMsg;
                    AppendLog(errorMsg, "WARN");
                    return;
                }

                var respLen = (ushort)(header[4] << 8 | header[5]);
                var body = new byte[respLen - 1];
                var offset = 0;
                while (offset < body.Length)
                {
                    var r = await stream.ReadAsync(body, offset, body.Length - offset);
                    if (r <= 0) break;
                    offset += r;
                }

                // 检查异常响应（功能码最高位为 1）
                bool isException = (body[0] & 0x80) != 0;
                var sb = new StringBuilder();
                sb.AppendLine($"功能码: 0x{fc:X2}  Unit={unitId}");
                sb.AppendLine($"起始地址: {start}");

                if (isException)
                {
                    byte exCode = body.Length > 1 ? body[1] : (byte)0;
                    string exDesc = exCode switch
                    {
                        0x01 => "非法功能码",
                        0x02 => "非法数据地址",
                        0x03 => "非法数据值",
                        0x04 => "从站设备故障",
                        _ => $"未知异常 (0x{exCode:X2})"
                    };
                    sb.AppendLine($"异常响应: {exDesc} (0x{exCode:X2})");
                }
                else
                {
                    sb.AppendLine("写入成功");
                    sb.AppendLine($"响应 (hex): {BitConverter.ToString(body)}");
                }

                string result = sb.ToString();
                if (respBox != null) respBox.Text = result;
                AppendLog($"Modbus 写入完成: {(isException ? "异常" : "成功")}");
                AppendDetailedLog("Modbus Write", host + ":" + port, null, isException ? "Exception" : "OK", result);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Modbus 写入错误: {ex.Message}";
                if (respBox != null) respBox.Text = errorMsg;
                AppendLog($"Modbus 写入错误: {ex.Message}", "ERROR");
            }
        }

        /// <summary>
        /// 原始 TCP 发送/接收的简易示例，发送文本并尝试读取少量响应。
        /// </summary>
        private async Task TestRawTcp(string host, string portText, string textToSend)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("Raw TCP：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!int.TryParse(portText, out var port))
            {
                MessageBox.Show("Raw TCP：端口无效", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            AppendLog($"Raw TCP -> {host}:{port} 发送长度 {textToSend?.Length ?? 0}");
            try
            {
                using var tcp = new TcpClient();
                var connectTask = tcp.ConnectAsync(host, port);
                if (await Task.WhenAny(connectTask, Task.Delay(4000)) != connectTask)
                {
                    string errorMsg = "Raw TCP: 连接超时";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _deviceProtoResponseBox.Text = "Raw TCP: 连接超时";
                    AppendLog("Raw TCP: 连接超时", "WARN");
                    return;
                }

                using var stream = tcp.GetStream();
                var sendBytes = Encoding.UTF8.GetBytes(textToSend ?? "");
                await stream.WriteAsync(sendBytes, 0, sendBytes.Length);

                var buffer = new byte[4096];
                var ms = new MemoryStream();
                stream.ReadTimeout = 3000;
                try
                {
                    var read = await stream.ReadAsync(buffer, 0, buffer.Length);
                    if (read > 0) ms.Write(buffer, 0, read);
                }
                catch { /* 读取超时或短连接时忽略 */ }

                var resp = Encoding.UTF8.GetString(ms.ToArray());
                string result = $"收到（{resp.Length} 字节）：{FirstChars(resp, 2000)}";
                MessageBox.Show(result, "Raw TCP 响应", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _deviceProtoResponseBox.Text = result;
                AppendLog("Raw TCP: 收到响应（文本）");
            }
            catch (Exception ex)
            {
                string errorMsg = $"Raw TCP 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _deviceProtoResponseBox.Text = $"Raw TCP 错误: {ex.Message}";
                AppendLog($"Raw TCP 错误: {ex.Message}", "ERROR");
            }
        }

        #endregion

        #region Device / Industrial protocol helpers

        /// <summary>
        /// 基于 S7.Net 的 PLC 读写尝试（通过反射兼容不同的 S7.Net 实现）。
        /// 该方法会尝试使用反射构造 Plc 对象并调用 Open/Read/Write/Close。
        /// 若未安装 S7.Net 库，会提示用户安装依赖包。
        /// </summary>
        private async Task TestS7ReadWrite(string host, string portText, string address, string writeValue)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("S7 测试：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!int.TryParse(portText, out var port)) port = 102;
            AppendLog($"S7 ReadWrite -> {host}:{port} address={address} write={writeValue}");

            try
            {
                // 通过反射兼容不同的 S7.Net 包名（S7NetPlus / S7netplus / S7.Net）
                var plcType = Type.GetType("S7.Net.Plc, S7NetPlus") ?? Type.GetType("S7.Net.Plc, S7netplus") ?? Type.GetType("S7.Net.Plc, S7.Net");
                if (plcType == null)
                {
                    string errorMsg = "未检测到 S7.Net 库。若需完整 S7 读写功能，请安装 S7NetPlus。";
                    MessageBox.Show(errorMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    AppendLog("未检测到 S7.Net 库。若需完整 S7 读写功能，请安装 S7NetPlus。", "WARN");
                    _deviceProtoResponseBox?.AppendText("未检测到 S7 库，无法执行读写。请安装 S7NetPlus。" + Environment.NewLine);
                    return;
                }

                // 寻找 CpuType（枚举）并尝试构造 Plc 对象（兼容多种构造函数签名）
                var cpuType = plcType.Assembly.GetType("S7.Net.CpuType") ?? plcType.Assembly.GetType("S7.Net.EnumTypes.CpuType");
                object cpu = null;
                if (cpuType != null && Enum.GetNames(cpuType).Contains("S71200")) cpu = Enum.Parse(cpuType, "S71200");
                else if (cpuType != null && Enum.GetNames(cpuType).Length > 0) cpu = Enum.GetValues(cpuType).GetValue(0);

                object plc = null;
                ConstructorInfo ctor = null;
                if (cpu != null)
                {
                    ctor = plcType.GetConstructor(new[] { cpuType, typeof(string), typeof(int), typeof(int) });
                    if (ctor != null) plc = ctor.Invoke(new object[] { cpu, host, 0, 1 });
                }

                if (plc == null)
                {
                    ctor = plcType.GetConstructor(new[] { typeof(string), typeof(int), typeof(int) });
                    if (ctor != null) plc = ctor.Invoke(new object[] { host, 0, 1 });
                }

                if (plc == null)
                {
                    string errorMsg = "无法构造 Plc 对象，请检查 S7NetPlus 版本。";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    AppendLog("无法构造 Plc 对象，请检查 S7NetPlus 版本。", "ERROR");
                    _deviceProtoResponseBox?.AppendText("无法构造 Plc 对象，请检查 S7NetPlus。" + Environment.NewLine);
                    return;
                }

                var open = plcType.GetMethod("Open");
                var close = plcType.GetMethod("Close");
                var read = plcType.GetMethod("Read", new[] { typeof(string) }) ?? plcType.GetMethod("ReadBytes", new[] { typeof(DataType), typeof(int), typeof(int) });
                var write = plcType.GetMethod("Write", new[] { typeof(string), typeof(object) }) ?? plcType.GetMethod("WriteBytes", new[] { typeof(DataType), typeof(int), typeof(byte[]) });

                try { open?.Invoke(plc, null); }
                catch (Exception ex)
                {
                    string errorMsg = $"打开 PLC 失败: {ex.Message}";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    AppendLog($"打开 PLC 失败: {ex.Message}", "ERROR");
                    _deviceProtoResponseBox?.AppendText("打开 PLC 失败: " + ex.Message + Environment.NewLine);
                }

                if (string.IsNullOrWhiteSpace(writeValue))
                {
                    if (read != null)
                    {
                        try
                        {
                            var val = read.Invoke(plc, new object[] { address });
                            string result = $"读取 {address} = {val}";
                            MessageBox.Show(result, "S7 读取结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            _deviceProtoResponseBox?.AppendText($"读取 {address} = {val}" + Environment.NewLine);
                            AppendLog($"S7: 读取 {address} = {FirstChars(val?.ToString() ?? "<null>", 200)}");
                        }
                        catch (Exception ex)
                        {
                            string errorMsg = $"读取失败: {ex.Message}";
                            MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            _deviceProtoResponseBox?.AppendText("读取失败: " + ex.Message + Environment.NewLine);
                            AppendLog($"S7 读取失败: {ex.Message}", "ERROR");
                        }
                    }
                    else
                    {
                        string errorMsg = "未找到适用的 Read 方法，请手动实现读取逻辑。";
                        MessageBox.Show(errorMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _deviceProtoResponseBox?.AppendText("未找到适用的 Read 方法，请手动实现读取逻辑。" + Environment.NewLine);
                    }
                }
                else
                {
                    if (write != null)
                    {
                        try
                        {
                            object toWrite = writeValue;
                            if (int.TryParse(writeValue, out var ival)) toWrite = ival;
                            else if (double.TryParse(writeValue, out var dval)) toWrite = dval;

                            write.Invoke(plc, new object[] { address, toWrite });
                            string result = $"写入成功 {address} = {writeValue}";
                            MessageBox.Show(result, "S7 写入结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            _deviceProtoResponseBox?.AppendText($"写入成功 {address} = {writeValue}" + Environment.NewLine);
                            AppendLog($"S7: 写入 {address} = {writeValue}");
                        }
                        catch (Exception ex)
                        {
                            string errorMsg = $"写入失败: {ex.Message}";
                            MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            _deviceProtoResponseBox?.AppendText("写入失败: " + ex.Message + Environment.NewLine);
                            AppendLog($"S7 写入失败: {ex.Message}", "ERROR");
                        }
                    }
                    else
                    {
                        string errorMsg = "未找到适用的 Write 方法，请手动实现写入逻辑。";
                        MessageBox.Show(errorMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _deviceProtoResponseBox?.AppendText("未找到适用的 Write 方法，请手动实现写入逻辑。" + Environment.NewLine);
                    }
                }

                try { close?.Invoke(plc, null); } catch { }
            }
            catch (Exception ex)
            {
                string errorMsg = $"S7 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"S7 错误: {ex.Message}", "ERROR");
                _deviceProtoResponseBox?.AppendText("错误: " + ex.Message + Environment.NewLine);
            }

            // 保持异步签名，方法体主要为同步反射调用
            await Task.Yield();
        }

        #endregion

        #region OPC UA and OPC DA helpers

        /// <summary>
        /// 通过检测已安装的客户端库（Opc.UaFx 或 OPC Foundation 官方库）尝试读取 OPC UA 节点。
        /// 对于不同库使用反射调用以避免编译时强依赖。
        /// </summary>
        private async Task DoOpcUaRead(string endpointUrl, string nodeId)
        {
            if (string.IsNullOrWhiteSpace(endpointUrl))
            {
                MessageBox.Show("OPC UA：Endpoint 为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(nodeId))
            {
                MessageBox.Show("OPC UA：NodeId 为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            AppendLog($"OPC UA -> {endpointUrl} 读取 {nodeId}");

            // 优先尝试 Opc.UaFx（更简单的 API），使用反射避免强耦合
            var uaFxType = Type.GetType("Opc.UaFx.Client.OpcClient, Opc.UaFx.Client");
            if (uaFxType != null)
            {
                try
                {
                    var client = Activator.CreateInstance(uaFxType, endpointUrl);
                    var connectMethod = uaFxType.GetMethod("Connect");
                    var readMethod = uaFxType.GetMethod("ReadNode", new[] { typeof(string) });
                    var disconnectMethod = uaFxType.GetMethod("Disconnect");

                    connectMethod?.Invoke(client, null);
                    string successMsg = "OPC UA: 已连接 (Opc.UaFx)";
                    MessageBox.Show(successMsg, "OPC UA 连接", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    AppendLog("OPC UA: 已连接 (Opc.UaFx)");

                    var value = readMethod?.Invoke(client, new object[] { nodeId });
                    string result = $"读取值: {value?.ToString() ?? "<null>"}";
                    MessageBox.Show(result, "OPC UA 读取结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    _opcUaRespBox.Text = value?.ToString() ?? "<null>";
                    AppendLog($"OPC UA: 读取值: {FirstChars(_opcUaRespBox.Text, 1000)}");

                    disconnectMethod?.Invoke(client, null);
                    AppendLog("OPC UA: 已断开");
                    return;
                }
                catch (TargetInvocationException tie)
                {
                    string errorMsg = $"OPC UA (Opc.UaFx) 内部异常: {tie.InnerException?.Message ?? tie.Message}";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    AppendLog($"OPC UA (Opc.UaFx) 内部异常: {tie.InnerException?.Message ?? tie.Message}", "ERROR");
                    _opcUaRespBox.Text = "错误: " + (tie.InnerException?.Message ?? tie.Message);
                    return;
                }
                catch (Exception ex)
                {
                    string errorMsg = $"OPC UA (Opc.UaFx) 错误: {ex.Message}";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    AppendLog($"OPC UA (Opc.UaFx) 错误: {ex.Message}", "ERROR");
                    _opcUaRespBox.Text = "错误: " + ex.Message;
                    return;
                }
            }

            // 如果存在官方 OPC UA 库，但调用过程较复杂，则给出提示
            var sessionType = Type.GetType("Opc.Ua.Client.Session, OPCFoundation.NetStandard.Opc.Ua");
            if (sessionType != null)
            {
                string infoMsg = "检测到 OPC Foundation .NET Standard 库，但示例调用较复杂。建议使用 Opc.UaFx.Client 以简化操作。";
                MessageBox.Show(infoMsg, "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                AppendLog("检测到 OPC Foundation .NET Standard 库，但示例调用较复杂。建议使用 Opc.UaFx.Client 以简化操作。", "INFO");
                _opcUaRespBox.Text = "检测到官方 OPC UA 库，但当前示例不自动调用。建议安装 Opc.UaFx.Client 并重试。";
                return;
            }

            string warningMsg = "未检测到 Opc.UaFx 客户端或官方 OPC UA 库。若需 OPC UA 支持，请安装 Opc.UaFx.Client";
            MessageBox.Show(warningMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            AppendLog("未检测到 Opc.UaFx 客户端或官方 OPC UA 库。若需 OPC UA 支持，请安装 Opc.UaFx.Client", "WARN");
            _opcUaRespBox.Text = "未检测到 OPC UA 客户端库。请安装 Opc.UaFx.Client 并重试。";
        }

        /// <summary>
        /// OPC UA 写节点示例（反射实现，优先使用 Opc.UaFx.Client）。
        /// </summary>
        private async Task DoOpcUaWrite(string endpointUrl, string nodeId, string value)
        {
            if (string.IsNullOrWhiteSpace(endpointUrl) || string.IsNullOrWhiteSpace(nodeId))
            {
                MessageBox.Show("OPC UA 写入：参数不足", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            AppendLog($"OPC UA 写入 -> {endpointUrl} {nodeId} = {value}");

            var uaFxType = Type.GetType("Opc.UaFx.Client.OpcClient, Opc.UaFx.Client");
            if (uaFxType != null)
            {
                try
                {
                    var client = Activator.CreateInstance(uaFxType, endpointUrl);
                    var connectMethod = uaFxType.GetMethod("Connect");
                    var writeMethod = uaFxType.GetMethod("WriteNode", new[] { typeof(string), typeof(object) }) ?? uaFxType.GetMethod("Write", new[] { typeof(string), typeof(object) });
                    var disconnectMethod = uaFxType.GetMethod("Disconnect");

                    connectMethod?.Invoke(client, null);
                    string successMsg = "OPC UA: 已连接 (Opc.UaFx) 写入...";
                    MessageBox.Show(successMsg, "OPC UA 连接", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    AppendLog("OPC UA: 已连接 (Opc.UaFx) 写入...");

                    writeMethod?.Invoke(client, new object[] { nodeId, value });
                    string result = $"写入 {nodeId} = {value}";
                    MessageBox.Show(result, "OPC UA 写入结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (_opcUaRespBox != null) _opcUaRespBox.Text = $"写入 {nodeId} = {value}";
                    AppendLog($"OPC UA: 写入 {nodeId} = {value}");

                    disconnectMethod?.Invoke(client, null);
                    AppendLog("OPC UA: 已断开");
                    return;
                }
                catch (Exception ex)
                {
                    string errorMsg = $"OPC UA 写错误: {ex.Message}";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    AppendLog($"OPC UA 写错误: {ex.Message}", "ERROR");
                    if (_opcUaRespBox != null) _opcUaRespBox.Text = "写错误: " + ex.Message;
                    return;
                }
            }

            var sessionType = Type.GetType("Opc.Ua.Client.Session, OPCFoundation.NetStandard.Opc.Ua");
            if (sessionType != null)
            {
                string infoMsg = "检测到官方 OPC UA 库，但当前示例不自动写入。建议使用 Opc.UaFx.Client 或添加官方库的写入实现。";
                MessageBox.Show(infoMsg, "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                AppendLog("检测到官方 OPC UA 库，但当前示例不自动写入。建议使用 Opc.UaFx.Client 或添加官方库的写入实现。", "INFO");
                if (_opcUaRespBox != null) _opcUaRespBox.Text = "检测到官方 OPC UA 库，但未实现写入示例。";
                return;
            }

            string warningMsg = "未检测到 OPC UA 客户端库，写入不可用。请安装 Opc.UaFx.Client。";
            MessageBox.Show(warningMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            AppendLog("未检测到 OPC UA 客户端库，写入不可用。请安装 Opc.UaFx.Client。", "WARN");
            if (_opcUaRespBox != null) _opcUaRespBox.Text = "未安装 Opc.UaFx.Client，无法写入。";
            await Task.Yield();
        }

        /// <summary>
        /// 使用 COM（OPC DA）进行读写操作（通过 ProgID 调用），仅适用于运行在支持 COM 的 Windows 环境。
        /// 使用后会尝试释放 COM 对象。
        /// </summary>
        private async Task DoOpcDaReadWrite(string host, string progId, string itemId, string writeValue)
        {
            if (string.IsNullOrWhiteSpace(progId))
            {
                MessageBox.Show("OPC DA: ProgID 为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                var comType = Type.GetTypeFromProgID(progId);
                if (comType == null)
                {
                    string errorMsg = $"ProgID 未注册: {progId}";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (_deviceProtoResponseBox != null) _deviceProtoResponseBox.Text = $"ProgID 未注册: {progId}";
                    AppendLog($"OPC DA ProgID 未注册: {progId}", "ERROR");
                    return;
                }

                dynamic server = Activator.CreateInstance(comType);
                try { server.Connect?.Invoke(progId); } catch { }

                if (!string.IsNullOrWhiteSpace(writeValue))
                {
                    try
                    {
                        server.Write?.Invoke(itemId, writeValue);
                        string result = $"尝试写入 {itemId} = {writeValue}";
                        MessageBox.Show(result, "OPC DA 写入", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (_deviceProtoResponseBox != null) _deviceProtoResponseBox.Text = $"尝试写入 {itemId} = {writeValue}";
                        AppendLog($"OPC DA: 写入尝试 {itemId} = {writeValue}");
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"写入尝试失败: {ex.Message}";
                        MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (_deviceProtoResponseBox != null) _deviceProtoResponseBox.Text = "写入尝试失败: " + ex.Message;
                        AppendLog($"OPC DA 写入失败: {ex.Message}", "ERROR");
                    }
                }
                else if (!string.IsNullOrWhiteSpace(itemId))
                {
                    try
                    {
                        var val = server.Read?.Invoke(itemId);
                        string result = $"读取 {itemId} = {val}";
                        MessageBox.Show(result, "OPC DA 读取", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (_deviceProtoResponseBox != null) _deviceProtoResponseBox.Text = $"读取 {itemId} = {val}";
                        AppendLog($"OPC DA: 读取 {itemId} = {val}");
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"读取尝试失败: {ex.Message}";
                        MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (_deviceProtoResponseBox != null) _deviceProtoResponseBox.Text = "读取尝试失败: " + ex.Message;
                        AppendLog($"OPC DA 读取失败: {ex.Message}", "ERROR");
                    }
                }

                try { Marshal.FinalReleaseComObject(server); } catch { }
            }
            catch (Exception ex)
            {
                string errorMsg = $"OPC DA 操作错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"OPC DA 操作错误: {ex.Message}", "ERROR");
                if (_deviceProtoResponseBox != null) _deviceProtoResponseBox.Text = "错误: " + ex.Message;
            }

            await Task.Yield();
        }

        #endregion

        #region Network Diagnostics (Ping / DNS / Port Scan / SSL Cert)

        /// <summary>
        /// Ping 目标主机，发送指定次数 ICMP 回显请求并统计结果。
        /// </summary>
        /// <param name="host">目标主机名或 IP</param>
        /// <param name="countText">Ping 次数（默认 4）</param>
        /// <param name="respBox">结果显示控件</param>
        private async Task DoPing(string host, string countText, TextBox respBox)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("Ping：目标为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!int.TryParse(countText, out var count) || count <= 0) count = 4;

            AppendLog($"Ping -> {host}，次数 {count}");
            try
            {
                using var ping = new Ping();
                var sb = new StringBuilder();
                int success = 0, fail = 0;
                long totalRtt = 0;

                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        var reply = await ping.SendPingAsync(host, 5000);
                        if (reply.Status == IPStatus.Success)
                        {
                            success++;
                            totalRtt += reply.RoundtripTime;
                            sb.AppendLine($"[{i + 1}] 回复: {reply.Address}  字节={reply.Buffer.Length}  时间={reply.RoundtripTime}ms  TTL={reply.Options?.Ttl ?? 0}");
                        }
                        else
                        {
                            fail++;
                            sb.AppendLine($"[{i + 1}] 失败: {reply.Status}");
                        }
                    }
                    catch (Exception ex)
                    {
                        fail++;
                        sb.AppendLine($"[{i + 1}] 异常: {ex.Message}");
                    }
                    if (i < count - 1) await Task.Delay(500);
                }

                sb.AppendLine();
                sb.AppendLine($"--- 统计 ---");
                sb.AppendLine($"发送={count}  接收={success}  丢失={fail}  丢包率={(count > 0 ? fail * 100.0 / count : 0):F1}%");
                if (success > 0) sb.AppendLine($"平均 RTT={totalRtt / success}ms");

                string result = sb.ToString();
                if (respBox != null) respBox.Text = result;
                AppendLog($"Ping 完成: 发送={count} 接收={success} 丢失={fail}");
                AppendDetailedLog("Ping", host, null, $"Sent={count} Rcv={success}", result);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Ping 错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"Ping 错误: {ex.Message}", "ERROR");
                if (respBox != null) respBox.Text = errorMsg;
            }
        }

        /// <summary>
        /// DNS 解析：查询域名的 A/AAAA/CNAME/MX 记录。
        /// </summary>
        /// <param name="domain">要解析的域名</param>
        /// <param name="respBox">结果显示控件</param>
        private async Task DoDnsLookup(string domain, TextBox respBox)
        {
            if (string.IsNullOrWhiteSpace(domain))
            {
                MessageBox.Show("DNS：域名为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            AppendLog($"DNS 解析 -> {domain}");
            try
            {
                var sb = new StringBuilder();
                var sw = Stopwatch.StartNew();

                // 基本解析
                var entry = await Dns.GetHostEntryAsync(domain);
                sw.Stop();

                sb.AppendLine($"主机名: {entry.HostName}");
                sb.AppendLine($"解析耗时: {sw.ElapsedMilliseconds}ms");
                sb.AppendLine();

                if (entry.AddressList.Length > 0)
                {
                    sb.AppendLine($"--- IP 地址 ({entry.AddressList.Length} 个) ---");
                    foreach (var ip in entry.AddressList)
                    {
                        string type = ip.AddressFamily == AddressFamily.InterNetwork ? "IPv4" :
                                      ip.AddressFamily == AddressFamily.InterNetworkV6 ? "IPv6" : ip.AddressFamily.ToString();
                        sb.AppendLine($"  [{type}] {ip}");
                    }
                }

                if (entry.Aliases.Length > 0)
                {
                    sb.AppendLine();
                    sb.AppendLine($"--- 别名 ({entry.Aliases.Length} 个) ---");
                    foreach (var alias in entry.Aliases) sb.AppendLine($"  {alias}");
                }

                string result = sb.ToString();
                if (respBox != null) respBox.Text = result;
                AppendLog($"DNS 解析完成: {domain} -> {entry.AddressList.Length} 个 IP");
                AppendDetailedLog("DNS Lookup", domain, null, $"Addresses={entry.AddressList.Length}", result);
            }
            catch (Exception ex)
            {
                string errorMsg = $"DNS 解析错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"DNS 解析错误: {ex.Message}", "ERROR");
                if (respBox != null) respBox.Text = errorMsg;
            }
        }

        // 端口扫描取消令牌（用于停止正在进行的扫描）
        private CancellationTokenSource _portScanCts;

        /// <summary>
        /// 端口扫描：对指定主机的端口范围进行 TCP 连接探测，带实时进度反馈。
        /// </summary>
        /// <param name="host">目标主机</param>
        /// <param name="startPortText">起始端口</param>
        /// <param name="endPortText">结束端口</param>
        /// <param name="respBox">结果显示控件</param>
        private async Task DoPortScan(string host, string startPortText, string endPortText, TextBox respBox)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("端口扫描：目标为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!int.TryParse(startPortText, out var startPort) || startPort < 1) startPort = 1;
            if (!int.TryParse(endPortText, out var endPort) || endPort < startPort) endPort = startPort + 100;
            if (endPort - startPort > 1024)
            {
                MessageBox.Show("端口扫描：范围过大（最多 1024 个端口），已自动截断", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                endPort = startPort + 1024;
            }

            // 取消上一次未完成的扫描
            _portScanCts?.Cancel();
            _portScanCts = new CancellationTokenSource();
            var cts = _portScanCts;

            int totalPorts = endPort - startPort + 1;
            AppendLog($"端口扫描 -> {host}:{startPort}-{endPort}（共 {totalPorts} 个端口）");

            // 常见端口映射
            string GetServiceName(int port) => port switch
            {
                21 => "FTP", 22 => "SSH", 23 => "Telnet", 25 => "SMTP",
                53 => "DNS", 80 => "HTTP", 110 => "POP3", 143 => "IMAP",
                443 => "HTTPS", 445 => "SMB", 993 => "IMAPS", 995 => "POP3S",
                1433 => "MSSQL", 1521 => "Oracle", 3306 => "MySQL", 3389 => "RDP",
                502 => "Modbus", 5432 => "PostgreSQL", 6379 => "Redis",
                8080 => "HTTP-Alt", 8443 => "HTTPS-Alt", 4840 => "OPC-UA",
                _ => ""
            };

            // 立即显示扫描状态
            if (respBox != null)
            {
                respBox.Text = $"正在扫描 {host}:{startPort}-{endPort}（共 {totalPorts} 个端口）...\r\n请稍候，扫描过程中会实时更新结果。\r\n";
            }

            var openPorts = new List<(int Port, string Service, long Ms)>();
            var totalSw = Stopwatch.StartNew();
            int scanned = 0;
            object lockObj = new();

            try
            {
                // 并发扫描，限制并发度为 80，超时 1.5 秒
                var semaphore = new SemaphoreSlim(80);
                var tasks = new List<Task>();

                for (int port = startPort; port <= endPort; port++)
                {
                    if (cts.Token.IsCancellationRequested) break;

                    int p = port;
                    tasks.Add(Task.Run(async () =>
                    {
                        await semaphore.WaitAsync(cts.Token);
                        try
                        {
                            using var tcp = new TcpClient();
                            var connectTask = tcp.ConnectAsync(host, p);
                            if (await Task.WhenAny(connectTask, Task.Delay(1500, cts.Token)) == connectTask && tcp.Connected)
                            {
                                long ms = 0;
                                lock (lockObj)
                                {
                                    openPorts.Add((p, GetServiceName(p), ms));
                                }
                            }
                        }
                        catch (OperationCanceledException) { }
                        catch { }
                        finally
                        {
                            semaphore.Release();
                            int done = Interlocked.Increment(ref scanned);

                            // 每扫描 100 个端口或发现开放端口时，更新一次 UI
                            if (done % 100 == 0 || done == totalPorts)
                            {
                                UpdatePortScanProgress(respBox, host, startPort, endPort, done, totalPorts, openPorts);
                            }
                        }
                    }, cts.Token));
                }

                await Task.WhenAll(tasks);
                totalSw.Stop();

                // 最终结果
                openPorts.Sort((a, b) => a.Port.CompareTo(b.Port));
                var sb = new StringBuilder();
                sb.AppendLine($"=== 端口扫描完成 ===");
                sb.AppendLine($"目标: {host}:{startPort}-{endPort}");
                sb.AppendLine($"扫描端口数: {totalPorts}");
                sb.AppendLine($"耗时: {totalSw.ElapsedMilliseconds}ms");
                sb.AppendLine();

                if (openPorts.Count > 0)
                {
                    sb.AppendLine($"--- 开放端口 ({openPorts.Count} 个) ---");
                    sb.AppendLine($"{"端口",-8} {"服务",-12}");
                    sb.AppendLine(new string('-', 22));
                    foreach (var (port, service, _) in openPorts)
                    {
                        sb.AppendLine($"{port,-8} {service,-12}");
                    }
                }
                else
                {
                    sb.AppendLine("--- 未发现开放端口 ---");
                }

                string result = sb.ToString();
                if (respBox != null) respBox.Text = result;
                AppendLog($"端口扫描完成: {openPorts.Count} 个开放端口，耗时 {totalSw.ElapsedMilliseconds}ms");
                AppendDetailedLog("Port Scan", host, null, $"Open={openPorts.Count}", result);
            }
            catch (OperationCanceledException)
            {
                if (respBox != null) respBox.AppendText("\r\n--- 扫描已取消 ---\r\n");
                AppendLog("端口扫描已取消", "WARN");
            }
            catch (Exception ex)
            {
                string errorMsg = $"端口扫描错误: {ex.Message}";
                if (respBox != null) respBox.Text = errorMsg;
                AppendLog($"端口扫描错误: {ex.Message}", "ERROR");
            }
        }

        /// <summary>
        /// 更新端口扫描进度到 UI（线程安全）。
        /// </summary>
        private void UpdatePortScanProgress(TextBox respBox, string host, int startPort, int endPort, int scanned, int total, List<(int Port, string Service, long Ms)> openPorts)
        {
            if (respBox == null) return;

            var sb = new StringBuilder();
            sb.AppendLine($"正在扫描 {host}:{startPort}-{endPort}...");
            sb.AppendLine($"进度: {scanned}/{total} ({scanned * 100 / total}%)");
            sb.AppendLine();

            if (openPorts.Count > 0)
            {
                sb.AppendLine($"--- 已发现 {openPorts.Count} 个开放端口 ---");
                lock (openPorts)
                {
                    foreach (var (port, service, _) in openPorts)
                    {
                        sb.AppendLine($"  {port,-8} {service}");
                    }
                }
            }

            if (this.InvokeRequired)
                this.BeginInvoke(() => respBox.Text = sb.ToString());
            else
                respBox.Text = sb.ToString();
        }

        /// <summary>
        /// 检查远程服务器的 SSL/TLS 证书信息。
        /// </summary>
        /// <param name="host">目标主机</param>
        /// <param name="portText">端口（默认 443）</param>
        /// <param name="respBox">结果显示控件</param>
        private async Task DoSslCertCheck(string host, string portText, TextBox respBox)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                MessageBox.Show("SSL 证书检查：主机为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!int.TryParse(portText, out var port) || port <= 0) port = 443;

            AppendLog($"SSL 证书检查 -> {host}:{port}");
            try
            {
                using var tcp = new TcpClient();
                var connectTask = tcp.ConnectAsync(host, port);
                if (await Task.WhenAny(connectTask, Task.Delay(5000)) != connectTask)
                {
                    string errorMsg = $"SSL 证书检查：TCP 连接超时 ({host}:{port})";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (respBox != null) respBox.Text = errorMsg;
                    AppendLog(errorMsg, "WARN");
                    return;
                }

                using var sslStream = new SslStream(tcp.GetStream(), false, (sender, cert, chain, errors) => true);
                var authTask = sslStream.AuthenticateAsClientAsync(host);
                if (await Task.WhenAny(authTask, Task.Delay(10000)) != authTask)
                {
                    string errorMsg = "SSL 握手超时";
                    MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (respBox != null) respBox.Text = errorMsg;
                    AppendLog(errorMsg, "WARN");
                    return;
                }

                var cert = sslStream.RemoteCertificate as X509Certificate2;
                if (cert == null)
                {
                    string msg = "未获取到证书信息";
                    MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (respBox != null) respBox.Text = msg;
                    return;
                }

                var sb = new StringBuilder();
                sb.AppendLine($"=== SSL/TLS 证书信息: {host}:{port} ===");
                sb.AppendLine();
                sb.AppendLine($"协议版本: {sslStream.SslProtocol}");
                sb.AppendLine($"加密算法: {sslStream.CipherAlgorithm} {sslStream.CipherStrength}bit");
                sb.AppendLine($"哈希算法: {sslStream.HashAlgorithm} {sslStream.HashStrength}bit");
                sb.AppendLine();
                sb.AppendLine($"主题: {cert.Subject}");
                sb.AppendLine($"颁发者: {cert.Issuer}");
                sb.AppendLine($"序列号: {cert.SerialNumber}");
                sb.AppendLine($"指纹 (SHA1): {cert.Thumbprint}");
                sb.AppendLine();
                sb.AppendLine($"生效时间: {cert.NotBefore:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"过期时间: {cert.NotAfter:yyyy-MM-dd HH:mm:ss}");

                var daysLeft = (cert.NotAfter - DateTime.Now).Days;
                string daysNote = daysLeft < 0 ? "【已过期！】" :
                                  daysLeft < 30 ? $"【即将过期，剩余 {daysLeft} 天】" :
                                  $"剩余 {daysLeft} 天";
                sb.AppendLine($"状态: {daysNote}");

                // SAN (Subject Alternative Names)
                var san = cert.Extensions["2.5.29.17"];
                if (san != null)
                {
                    sb.AppendLine();
                    sb.AppendLine($"SAN (主题备用名称):");
                    sb.AppendLine($"  {san.Format(true).Trim()}");
                }

                // 证书链
                using var chain = new X509Chain();
                chain.Build(cert);
                if (chain.ChainElements.Count > 1)
                {
                    sb.AppendLine();
                    sb.AppendLine($"--- 证书链 ({chain.ChainElements.Count} 级) ---");
                    for (int i = 0; i < chain.ChainElements.Count; i++)
                    {
                        var elem = chain.ChainElements[i];
                        sb.AppendLine($"  [{i}] {elem.Certificate.Subject}");
                        if (elem.ChainElementStatus.Length > 0)
                        {
                            foreach (var status in elem.ChainElementStatus)
                                sb.AppendLine($"      状态: {status.Status} - {status.StatusInformation}");
                        }
                    }
                }

                string result = sb.ToString();
                if (respBox != null) respBox.Text = result;
                AppendLog($"SSL 证书检查完成: {cert.Subject}, 有效期至 {cert.NotAfter:yyyy-MM-dd}, {daysNote}");
                AppendDetailedLog("SSL Cert Check", host + ":" + port, null, $"Valid until {cert.NotAfter:yyyy-MM-dd}", result);
            }
            catch (Exception ex)
            {
                string errorMsg = $"SSL 证书检查错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"SSL 证书检查错误: {ex.Message}", "ERROR");
                if (respBox != null) respBox.Text = errorMsg;
            }
        }

        #endregion

        /// <summary>
        /// 截取字符串前若干字符并添加省略号（用于日志摘要展示）。
        /// </summary>
        private string FirstChars(string text, int maxLen)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Length <= maxLen ? text : text.Substring(0, maxLen) + "...";
        }

        /// <summary>
        /// 打开日志目录（在文件资源管理器中）。
        /// </summary>
        private void OpenLogFolder()
        {
            try
            {
                System.Diagnostics.Process.Start("explorer.exe", _logDir);
            }
            catch (Exception ex)
            {
                string errorMsg = $"打开日志文件夹错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"打开日志文件夹错误: {ex.Message}", "ERROR");
            }
        }

        /// <summary>
        /// 将当天日志目录压缩为 zip 并提示保存位置（简单实现）。
        /// 注意：若日志目录很大，该操作可能耗时且占用磁盘空间。
        /// </summary>
        private void ExportTodayLog()
        {
            try
            {
                var date = DateTime.Now.ToString("yyyyMMdd");
                var zipPath = Path.Combine(_logDir, $"Log_{date}.zip");

                System.IO.Compression.ZipFile.CreateFromDirectory(_logDir, zipPath);

                string successMsg = $"导出当天日志成功: {zipPath}";
                MessageBox.Show(successMsg, "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                AppendLog($"导出当天日志成功: {zipPath}");
            }
            catch (Exception ex)
            {
                string errorMsg = $"导出日志错误: {ex.Message}";
                MessageBox.Show(errorMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"导出日志错误: {ex.Message}", "ERROR");
            }
        }
    }
}