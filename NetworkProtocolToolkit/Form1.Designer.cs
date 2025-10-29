using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;

namespace NetworkProtocolToolkit
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being be used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(1000, 820);
            Text = "网络协议工具箱 Ver 1.003";

            // Main TabControl with all pages.
            var tab = new TabControl { Dock = DockStyle.Fill };

            // Add all tabs
            tab.TabPages.Add(CreateTab("连通性检测", CreateConnectivityTab));
            tab.TabPages.Add(CreateTab("REST", CreateRestTab));
            tab.TabPages.Add(CreateTab("HTTP/HTTPS", CreateHttpTab));
            tab.TabPages.Add(CreateTab("WebSocket", CreateWebSocketTab));
            tab.TabPages.Add(CreateTab("WebService (SOAP)", CreateWebServiceTab));
            tab.TabPages.Add(CreateTab("FTP / SFTP", CreateFtpSftpTab));
            tab.TabPages.Add(CreateTab("SMTP", CreateSmtpTab));
            tab.TabPages.Add(CreateTab("上位机协议测试", CreateDeviceProtocolsTab));
            tab.TabPages.Add(CreateTab("收取邮件 (MailKit)", CreateMailReceiveTab));
            tab.TabPages.Add(CreateTab("日志", CreateLogTab));
            tab.TabPages.Add(CreateTab("帮助", CreateHelpTab));
            tab.TabPages.Add(CreateTab("更新日志", CreateUpdateLogTab));

            Controls.Add(tab);
        }

        private TabPage CreateTab(string title, Func<Control> contentGenerator)
        {
            // 支持两种工厂返回值：
            // - 如果返回 TabPage，直接使用（并在其 Text 为空时应用 title）
            // - 如果返回 Control，直接将该控件添加到 TabPage 并 Dock=Fill（避免嵌套 FlowLayoutPanel 导致布局被压缩）
            var content = contentGenerator();

            if (content is TabPage existingPage)
            {
                if (string.IsNullOrWhiteSpace(existingPage.Text)) existingPage.Text = title;
                return existingPage;
            }

            var page = new TabPage(title);

            // 直接把控件放入 TabPage，并让其填充，这能保证像 CreateConnectivityTab / CreateRestTab 返回的 FlowLayoutPanel 按预期显示。
            content.Dock = DockStyle.Fill;
            page.Controls.Add(content);
            return page;
        }

        private Control CreateConnectivityTab()
        {
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };

            // IP+Port 检测
            var ip = LabeledTextBox("目标 IP / 主机：", "127.0.0.1");
            var port = LabeledTextBox("目标端口：", "3306");
            var btnIpTest = new Button { Text = "检测 IP:Port 可达性", Width = 200 };
            btnIpTest.Click += async (_, __) =>
            {
                var result = await TestIpPort(ip.TextBox.Text, port.TextBox.Text);
            };

            panel.Controls.Add(new Label { Text = "—— IP + Port 通断检测 ——", Width = 940 });
            panel.Controls.Add(ip.Panel);
            panel.Controls.Add(port.Panel);
            panel.Controls.Add(btnIpTest);

            // 数据库检测区：通用字段 + 单独按钮
            panel.Controls.Add(new Label { Text = "—— 数据库连接检测 ——", Width = 940 });

            var dbServer = LabeledTextBox("数据库主机：", "localhost");
            var dbPort = LabeledTextBox("数据库端口（可留空使用默认）：", "");
            var dbName = LabeledTextBox("数据库名 / 服务名：", "testdb");
            var dbUser = LabeledTextBox("用户名：", "");
            var dbPass = LabeledTextBox("密码：", "");
            var dbConnStr = LabeledTextBox("完整连接字符串（可选，优先使用）：", "", multiline: true, height: 60);

            var btnTestSqlServer = new Button { Text = "测试 SQL Server", Width = 200 };
            btnTestSqlServer.Click += async (_, __) =>
            {
                var cs = string.IsNullOrWhiteSpace(dbConnStr.TextBox.Text)
                    ? BuildSqlServerConnectionString(dbServer.TextBox.Text, dbPort.TextBox.Text, dbName.TextBox.Text, dbUser.TextBox.Text, dbPass.TextBox.Text)
                    : dbConnStr.TextBox.Text;
                await TestDbConnection("sqlserver", cs);
            };

            var btnTestOracle = new Button { Text = "测试 Oracle", Width = 200 };
            btnTestOracle.Click += async (_, __) =>
            {
                var cs = string.IsNullOrWhiteSpace(dbConnStr.TextBox.Text)
                    ? BuildOracleConnectionString(dbServer.TextBox.Text, dbPort.TextBox.Text, dbName.TextBox.Text, dbUser.TextBox.Text, dbPass.TextBox.Text)
                    : dbConnStr.TextBox.Text;
                await TestDbConnection("oracle", cs);
            };

            var btnTestMySql = new Button { Text = "测试 MySQL", Width = 200 };
            btnTestMySql.Click += async (_, __) =>
            {
                var cs = string.IsNullOrWhiteSpace(dbConnStr.TextBox.Text)
                    ? BuildMySqlConnectionString(dbServer.TextBox.Text, dbPort.TextBox.Text, dbName.TextBox.Text, dbUser.TextBox.Text, dbPass.TextBox.Text)
                    : dbConnStr.TextBox.Text;
                await TestDbConnection("mysql", cs);
            };

            panel.Controls.Add(dbServer.Panel);
            panel.Controls.Add(dbPort.Panel);
            panel.Controls.Add(dbName.Panel);
            panel.Controls.Add(dbUser.Panel);
            panel.Controls.Add(dbPass.Panel);
            panel.Controls.Add(dbConnStr.Panel);

            var btnPanel = new FlowLayoutPanel { Width = 940, Height = 40 };
            btnPanel.Controls.Add(btnTestSqlServer);
            btnPanel.Controls.Add(btnTestOracle);
            btnPanel.Controls.Add(btnTestMySql);
            panel.Controls.Add(btnPanel);

            // 新增结果表格显示
            var dgvPanel = new Panel { Width = 940, Height = 240 };
            _dbResultGrid = new DataGridView
            {
                Left = 0,
                Top = 0,
                Width = 940,
                Height = 240,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnCount = 6
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
            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false
            };

            var url = LabeledTextBox("接口 URL（POST）：", "https://postman-echo.com/post");
            var json = LabeledTextBox("JSON 请求体：", "{\"a\":1}", multiline: true, height: 120);
            var btn = new Button { Text = "POST JSON 并显示响应", Width = 260 };
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
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };

            var urlBox = LabeledTextBox("URL（http/https）：", "https://postman-echo.com/get");
            var btnGet = new Button { Text = "GET 请求", Width = 120 };
            btnGet.Click += async (_, __) => await DoHttpGet(urlBox.TextBox.Text);

            var postBody = LabeledTextBox("POST JSON（请求体）：", "{\"hello\":\"world\"}", multiline: true, height: 100);
            var btnPost = new Button { Text = "POST 请求（application/json）", Width = 260 };
            btnPost.Click += async (_, __) => await DoHttpPost(urlBox.TextBox.Text, postBody.TextBox.Text);

            // 响应显示区
            var respPanel = new Panel { Width = 940, Height = 300 };
            var respLabel = new Label { Text = "响应（Body）:", Left = 0, Top = 4, Width = 200 };
            _httpResponseBox = new TextBox { Left = 210, Top = 0, Width = 700, Height = 290, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            respPanel.Controls.Add(respLabel);
            respPanel.Controls.Add(_httpResponseBox);

            panel.Controls.Add(urlBox.Panel);
            panel.Controls.Add(btnGet);
            panel.Controls.Add(btnPost);
            panel.Controls.Add(postBody.Panel);
            panel.Controls.Add(respPanel);

            page.Controls.Add(panel);
            return page;
        }

        private Control CreateWebSocketTab()
        {
            var page = new TabPage("WebSocket");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };

            var urlBox = LabeledTextBox("WebSocket 地址：", "wss://echo.websocket.org");
            var btnConnect = new Button { Text = "连接并发送 'hello'", Width = 200 };
            btnConnect.Click += async (_, __) => await DoWebSocketEcho(urlBox.TextBox.Text);

            panel.Controls.Add(urlBox.Panel);
            panel.Controls.Add(btnConnect);
            page.Controls.Add(panel);
            return page;
        }

        private Control CreateWebServiceTab()
        {
            var page = new TabPage("WebService (SOAP)");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };

            var url = LabeledTextBox("服务 URL：", "https://www.example.com/Service.svc");
            var soapAction = LabeledTextBox("SOAPAction（可选）：", "");
            var reqXml = LabeledTextBox("请求 XML：", "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\">\n  <soapenv:Body>\n    <!-- your request here -->\n  </soapenv:Body>\n</soapenv:Envelope>", multiline: true, height: 200);

            var btn = new Button { Text = "调用 WebService (SOAP)", Width = 260 };
            var respPanel = new Panel { Width = 940, Height = 320 };
            var respLabel = new Label { Text = "响应（Body）:", Left = 0, Top = 4, Width = 200 };
            _wsResponseBox = new TextBox { Left = 210, Top = 0, Width = 700, Height = 300, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            respPanel.Controls.Add(respLabel);
            respPanel.Controls.Add(_wsResponseBox);

            btn.Click += async (_, __) => await DoWebServiceTest(url.TextBox.Text, soapAction.TextBox.Text, reqXml.TextBox.Text);

            panel.Controls.Add(url.Panel);
            panel.Controls.Add(soapAction.Panel);
            panel.Controls.Add(reqXml.Panel);

            // Security note for WebService
            panel.Controls.Add(new Label { Text = "注意：如果使用 HTTPS，可能需要受信任的证书；某些服务要求额外认证（如 Basic、WS-Security）。", Width = 940, ForeColor = System.Drawing.Color.DarkRed, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, FontStyle.Italic) });

            panel.Controls.Add(btn);
            panel.Controls.Add(respPanel);

            page.Controls.Add(panel);
            return page;
        }

        private Control CreateFtpSftpTab()
        {
            var page = new TabPage("FTP / SFTP");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };

            var host = LabeledTextBox("主机（Host）：", "ftp.example.com");
            var port = LabeledTextBox("端口（Port，可留空使用默认）：", "");
            var user = LabeledTextBox("用户名：", "anonymous");
            var pass = LabeledTextBox("密码：", "");
            var keyPath = LabeledTextBox("私钥路径（SFTP 可选）：", "", multiline: false);
            var btnFtpList = new Button { Text = "FTP：列出根目录", Width = 200 };
            btnFtpList.Click += async (_, __) => await DoFtpList(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text);

            var btnSftpList = new Button { Text = "SFTP：列出根目录（Renci.SshNet）", Width = 320 };
            btnSftpList.Click += async (_, __) => await DoSftpList_Strong(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text, keyPath.TextBox.Text);

            panel.Controls.Add(host.Panel);
            panel.Controls.Add(port.Panel);
            panel.Controls.Add(user.Panel);
            panel.Controls.Add(pass.Panel);
            panel.Controls.Add(keyPath.Panel);
            panel.Controls.Add(btnFtpList);
            panel.Controls.Add(btnSftpList);

            // Security / anonymous note for FTP/SFTP
            panel.Controls.Add(new Label { Text = "注意：FTP 可支持匿名访问（anonymous），但很多服务器禁止匿名；SFTP 使用 SSH 身份验证（密码或私钥）。防火墙、被动/主动模式和 FTPS（TLS）都会影响连接。", Width = 940, ForeColor = System.Drawing.Color.DarkRed, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, FontStyle.Italic) });

            page.Controls.Add(panel);
            return page;
        }

        private Control CreateSmtpTab()
        {
            var page = new TabPage("SMTP");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };

            var host = LabeledTextBox("SMTP 主机：", "smtp.example.com");
            var port = LabeledTextBox("端口：", "25");
            var user = LabeledTextBox("用户名：", "");
            var pass = LabeledTextBox("密码：", "");
            var from = LabeledTextBox("发件人（From）：", "sender@example.com");
            var to = LabeledTextBox("收件人（To）：", "recipient@example.com");
            var subj = LabeledTextBox("主题：", "测试邮件");
            var body = LabeledTextBox("正文：", "来自协议调试器的测试邮件", multiline: true, height: 100);

            // 统一样式：使用与“收取邮件 (MailKit)”页相同的安全连接控件布局
            var sslPanel = new Panel { Width = 940, Height = 30 };
            var sslCheck = new CheckBox { Left = 210, Top = 4, Width = 200, Checked = false, Text = "使用 SSL/TLS" };
            var sslLabel = new Label { Left = 0, Top = 4, Width = 200, Text = "安全连接：" };
            sslPanel.Controls.Add(sslLabel);
            sslPanel.Controls.Add(sslCheck);

            // Security note for SMTP
            panel.Controls.Add(new Label { Text = "注意：大多数 SMTP 服务器要求身份验证和 TLS/SSL。常见端口：25 (可 STARTTLS)、465 (SSL)、587 (提交)。匿名发送通常被拒绝。", Width = 940, ForeColor = System.Drawing.Color.DarkRed, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, FontStyle.Italic) });

            var btnSend = new Button { Text = "发送邮件（SmtpClient）", Width = 220 };
            btnSend.Click += async (_, __) => await DoSendSmtp(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text, from.TextBox.Text, to.TextBox.Text, subj.TextBox.Text, body.TextBox.Text, sslCheck.Checked);

            // 保持控件顺序与收取邮件页一致以统一风格
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
            // inner tab control for protocols
            var innerTabs = new TabControl { Dock = DockStyle.Fill };

            // Use a TableLayoutPanel so the warning row always reserves space and cannot overlap the tab headers
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var topPanel = new Panel { Dock = DockStyle.Fill };
            var warningLabel = new Label
            {
                Text = "注意：设备协议通常在受控网络环境中，需要设备侧账号、访问控制或特殊路由。防火墙、端口和协议版本（例如 S7、Modbus 变体）会影响可用性。",
                Dock = DockStyle.Fill,
                ForeColor = System.Drawing.Color.DarkRed,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9, FontStyle.Italic),
                Padding = new Padding(6)
            };
            topPanel.Controls.Add(warningLabel);

            // Add top row and inner tabs into layout
            layout.Controls.Add(topPanel, 0, 0);
            innerTabs.Dock = DockStyle.Fill;
            layout.Controls.Add(innerTabs, 0, 1);

            // S7 tab
            var s7Page = new TabPage("S7");
            var s7Panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };
            var s7Host = LabeledTextBox("目标主机：", "192.168.1.100");
            var s7Port = LabeledTextBox("端口：", "102");
            var s7Address = LabeledTextBox("读写地址（例如 DB1,DBW0 或 P#DB1.DBD0）:", "DB1.DBW0");
            var s7WriteVal = LabeledTextBox("写入值：", "123");
            var s7ReadBtn = new Button { Text = "读 S7 点位", Width = 140 };
            var s7WriteBtn = new Button { Text = "写 S7 点位", Width = 140 };
            var s7Resp = new TextBox { Left = 210, Top = 0, Width = 680, Height = 120, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };

            s7ReadBtn.Click += async (_, __) =>
            {
                _deviceProtoResponseBox = s7Resp;
                await TestS7ReadWrite(s7Host.TextBox.Text, s7Port.TextBox.Text, s7Address.TextBox.Text, null);
            };
            s7WriteBtn.Click += async (_, __) =>
            {
                _deviceProtoResponseBox = s7Resp;
                await TestS7ReadWrite(s7Host.TextBox.Text, s7Port.TextBox.Text, s7Address.TextBox.Text, s7WriteVal.TextBox.Text);
            };

            s7Panel.Controls.Add(s7Host.Panel);
            s7Panel.Controls.Add(s7Port.Panel);
            s7Panel.Controls.Add(s7Address.Panel);
            s7Panel.Controls.Add(s7WriteVal.Panel);
            s7Panel.Controls.Add(s7ReadBtn);
            s7Panel.Controls.Add(s7WriteBtn);
            s7Panel.Controls.Add(s7Resp);
            s7Page.Controls.Add(s7Panel);
            innerTabs.TabPages.Add(s7Page);

            // Modbus tab
            var modbusPage = new TabPage("Modbus TCP");
            var modbusPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };
            var mbHost = LabeledTextBox("目标主机：", "192.168.1.100");
            var mbPort = LabeledTextBox("端口：", "502");
            var mbUnit = LabeledTextBox("UnitId：", "1");
            var mbStart = LabeledTextBox("起始寄存器：", "0");
            var mbQty = LabeledTextBox("读取数量：", "1");
            var mbBtn = new Button { Text = "读取寄存器", Width = 200 };
            var mbResp = new TextBox { Left = 210, Top = 0, Width = 680, Height = 120, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            mbBtn.Click += async (_, __) => { _deviceProtoResponseBox = mbResp; await TestModbusTcp(mbHost.TextBox.Text, mbPort.TextBox.Text, mbUnit.TextBox.Text, mbStart.TextBox.Text, mbQty.TextBox.Text); };
            modbusPanel.Controls.Add(mbHost.Panel);
            modbusPanel.Controls.Add(mbPort.Panel);
            modbusPanel.Controls.Add(mbUnit.Panel);
            modbusPanel.Controls.Add(mbStart.Panel);
            modbusPanel.Controls.Add(mbQty.Panel);
            modbusPanel.Controls.Add(mbBtn);
            modbusPanel.Controls.Add(mbResp);
            modbusPage.Controls.Add(modbusPanel);
            innerTabs.TabPages.Add(modbusPage);

            // OPC tab (integrated)
            var opcPage = new TabPage("OPC");
            var opcPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };
            // OPC UA area
            var uaEndpoint = LabeledTextBox("Endpoint URL:", "opc.tcp://localhost:4840");
            var uaNode = LabeledTextBox("NodeId:", "ns=2;s=Demo.Static.Scalar.Int32");
            var uaReadBtn = new Button { Text = "OPC UA: 读节点", Width = 140 };
            var uaWriteBtn = new Button { Text = "OPC UA: 写节点", Width = 140 };
            var uaWriteVal = LabeledTextBox("写入值：", "123");
            var uaResp = new TextBox { Left = 210, Top = 0, Width = 680, Height = 120, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };

            uaReadBtn.Click += async (_, __) => { _deviceProtoResponseBox = uaResp; await DoOpcUaRead(uaEndpoint.TextBox.Text, uaNode.TextBox.Text); };
            uaWriteBtn.Click += async (_, __) => { _deviceProtoResponseBox = uaResp; await DoOpcUaWrite(uaEndpoint.TextBox.Text, uaNode.TextBox.Text, uaWriteVal.TextBox.Text); };

            opcPanel.Controls.Add(uaEndpoint.Panel);
            opcPanel.Controls.Add(uaNode.Panel);
            opcPanel.Controls.Add(uaWriteVal.Panel);
            opcPanel.Controls.Add(uaReadBtn);
            opcPanel.Controls.Add(uaWriteBtn);
            opcPanel.Controls.Add(uaResp);

            // OPC DA area
            opcPanel.Controls.Add(new Label { Text = "OPC DA（本机 COM）", Width = 900 });
            var daProg = LabeledTextBox("ProgID:", "OPC.Automation.1");
            var daItemId = LabeledTextBox("ItemID (Item 名称):", "");
            var daReadBtn = new Button { Text = "OPC DA: 读项", Width = 140 };
            var daWriteBtn = new Button { Text = "OPC DA: 写项", Width = 140 };
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

            // Raw TCP tab
            var rawPage = new TabPage("Raw TCP");
            var rawPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };
            var rawHost = LabeledTextBox("目标主机：", "192.168.1.100");
            var rawPort = LabeledTextBox("端口：", "502");
            var rawPayload = LabeledTextBox("发送内容：", "Hello", multiline: true, height: 80);
            var rawBtn = new Button { Text = "发送并接收", Width = 200 };
            var rawResp = new TextBox { Left = 210, Top = 0, Width = 680, Height = 120, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
            rawBtn.Click += async (_, __) => { _deviceProtoResponseBox = rawResp; await TestRawTcp(rawHost.TextBox.Text, rawPort.TextBox.Text, rawPayload.TextBox.Text); };
            rawPanel.Controls.Add(rawHost.Panel);
            rawPanel.Controls.Add(rawPort.Panel);
            rawPanel.Controls.Add(rawPayload.Panel);
            rawPanel.Controls.Add(rawBtn);
            rawPanel.Controls.Add(rawResp);
            rawPage.Controls.Add(rawPanel);
            innerTabs.TabPages.Add(rawPage);

            layout.Controls.Add(innerTabs, 0, 1);

            page.Controls.Add(layout);
            return page;
        }

        private Control CreateMailReceiveTab()
        {
            var page = new TabPage("收取邮件 (MailKit)");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };

            var host = LabeledTextBox("邮件服务器主机：", "pop.example.com");
            var port = LabeledTextBox("端口：", "995");
            var user = LabeledTextBox("用户名：", "");
            var pass = LabeledTextBox("密码：", "");
            var sslPanel = new Panel { Width = 940, Height = 30 };
            var sslCheck = new CheckBox { Left = 210, Top = 4, Width = 200, Checked = true, Text = "使用 SSL/TLS" };
            var sslLabel = new Label { Left = 0, Top = 4, Width = 200, Text = "安全连接：" };
            sslPanel.Controls.Add(sslLabel);
            sslPanel.Controls.Add(sslCheck);

            var pop3Btn = new Button { Text = "POP3：列出前 5 封（MailKit）", Width = 260 };
            pop3Btn.Click += async (_, __) => await DoPop3MailKit(host.TextBox.Text, port.TextBox.Text, user.TextBox.Text, pass.TextBox.Text, sslCheck.Checked);

            var imapHost = LabeledTextBox("IMAP 主机（可与 POP3 不同）：", "imap.example.com");
            var imapPort = LabeledTextBox("IMAP 端口：", "993");
            var imapUser = LabeledTextBox("IMAP 用户名：", "");
            var imapPass = LabeledTextBox("IMAP 密码：", "");
            var imapSslPanel = new Panel { Width = 940, Height = 30 };
            var imapSslCheck = new CheckBox { Left = 210, Top = 4, Width = 200, Checked = true, Text = "使用 SSL/TLS" };
            var imapSslLabel = new Label { Left = 0, Top = 4, Width = 200, Text = "IMAP 安全：" };
            imapSslPanel.Controls.Add(imapSslLabel);
            imapSslPanel.Controls.Add(imapSslCheck);

            var imapBtn = new Button { Text = "IMAP：列出收件箱前 10 封（MailKit）", Width = 320 };
            imapBtn.Click += async (_, __) => await DoImapMailKit(imapHost.TextBox.Text, imapPort.TextBox.Text, imapUser.TextBox.Text, imapPass.TextBox.Text, imapSslCheck.Checked);

            panel.Controls.Add(host.Panel);
            panel.Controls.Add(port.Panel);
            panel.Controls.Add(user.Panel);
            panel.Controls.Add(pass.Panel);
            panel.Controls.Add(sslPanel);
            panel.Controls.Add(pop3Btn);

            // Security note for mail receive
            panel.Controls.Add(new Label { Text = "注意：POP3/IMAP 服务通常要求 TLS/SSL 和账号认证；部分大厂使用 OAuth2 登录（无法直接使用密码），请根据服务文档配置。", Width = 940, ForeColor = System.Drawing.Color.DarkRed, Font = new System.Drawing.Font("Microsoft Sans Serif", 9, FontStyle.Italic) });

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
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };

            _logBox = new TextBox { Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10), Width = 940, Height = 600 };
            var btns = new FlowLayoutPanel { Width = 940, Height = 30 };
            var btnClear = new Button { Text = "清除日志", Width = 100 };
            btnClear.Click += (_, __) => { lock (_logLock) { _logLines.Clear(); _logBox.Clear(); } };
            var btnExport = new Button { Text = "导出日志到文件", Width = 120 };
            btnExport.Click += (_, __) =>
            {
                try
                {
                    var dlg = new SaveFileDialog { Filter = "文本文件|*.txt", FileName = $"protocol-debug-log-{DateTime.Now:yyyyMMddHHmmss}.txt" };
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        File.WriteAllLines(dlg.FileName, _logLines);
                        MessageBox.Show("导出成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            var btnOpenFolder = new Button { Text = "打开日志文件夹", Width = 140 };
            btnOpenFolder.Click += (_, __) => OpenLogFolder();

            var btnExportToday = new Button { Text = "导出今日日志", Width = 120 };
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
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };

            var helpLabel = new Label { Text = "使用教程", Font = new System.Drawing.Font("Microsoft Sans Serif", 11, FontStyle.Bold), Width = 940, Margin = new Padding(0, 10, 0, 5) };

            var helpText = @"快速入门：

1) 连通性检测
 - 在“连通性检测”页填写目标主机和端口，点击“检测 IP:Port 可达性”测试 TCP 级别连通性。

2) REST / HTTP / WebService / WebSocket
 - 在对应页填写 URL（或请求体），点击对应按钮发送请求并查看响应。
 - REST 页默认发送 POST JSON；HTTP 页支持 GET 和 POST。
 - WebService 页可直接贴入 SOAP XML 并填写 SOAPAction（如需要）。
 - WebSocket 页适合测试 ws/wss 回显服务。

3) FTP / SFTP
 - 填写主机、端口、用户名、密码，FTP 列表仅适用于普通 FTP；SFTP 使用 Renci.SshNet（支持私钥）。

4) SMTP（发送邮件）
 - 填写 SMTP 主机、端口、用户名、密码、发件人、收件人、主题和正文。
 - 使用“安全连接”复选框选择是否启用 SSL/TLS（与收取邮件页风格一致）。

5) 收取邮件（POP3 / IMAP，MailKit）
 - 填写邮件服务器信息并选择“安全连接”启动 TLS/SSL，点击相应按钮列出邮件摘要。

6) 上位机协议测试（S7、Modbus、OPC、Raw TCP）
 - 选择对应子页，填写目标主机/端口/地址或节点，点击读写按钮进行测试。

7) 日志与问题定位
 - 所有操作会记录到“日志”页，出现问题请先查看当天日志并发送给作者以便排查。

使用建议：
 - 若遇到证书或握手失败，请检查目标服务是否使用自签名证书并在测试环境放通或替换为受信任证书。
 - 对于设备协议（S7/Modbus/OPC），确保网络、端口和防火墙规则允许访问。

示例：测试 SMTP TLS
 - SMTP 主机: smtp.example.com
 - 端口: 465（或 587）
 - 勾选 “使用 SSL/TLS”，填写用户名/密码，点击发送。";

            var helpBox = new TextBox { Left = 0, Top = 0, Width = 940, Height = 360, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, Font = new System.Drawing.Font("Consolas", 10), Text = helpText };

            // 作者的话
            var authorLabel = new Label { Text = "作者的话", Font = new System.Drawing.Font("Microsoft Sans Serif", 11, FontStyle.Bold), Width = 940, Margin = new Padding(0, 20, 0, 5) };
            var authorText = @"大家好！这是我根据前两年工作顺手整理的一份协议连通测试程序，把我觉得可能用到的都放进来了，也算是我个人工作的一点小收获吧，希望能帮到有需要的朋友！
不过，里面很多协议我目前条件有限，没法全部测试。所以如果大家测试时发现任何问题，麻烦把当天的日志发到我的邮箱：heyu8888888888@163.com。先谢谢啦！";
            var authorBox = new TextBox { Left = 0, Top = 0, Width = 940, Height = 180, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, Font = new System.Drawing.Font("Consolas", 10), Text = authorText };

            panel.Controls.Add(helpLabel);
            panel.Controls.Add(helpBox);
            panel.Controls.Add(authorLabel);
            panel.Controls.Add(authorBox);

            page.Controls.Add(panel);
            return page;
        }

        private Control CreateUpdateLogTab()
        {
            var page = new TabPage("更新日志");
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };

            // 更新日志内容
            var updateLogText = @"版本 1.003 - 2025年10月
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

            var updateLogLabel = new Label { Text = "更新日志", Font = new System.Drawing.Font("Microsoft Sans Serif", 11, FontStyle.Bold), Width = 940, Margin = new Padding(0, 10, 0, 5) };
            var updateLogBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 11),
                Width = 940,
                Height = 460,
                Text = updateLogText
            };

            panel.Controls.Add(updateLogLabel);
            panel.Controls.Add(updateLogBox);

            page.Controls.Add(panel);
            return page;
        }

        // helper to create labeled textbox panels
        private (Panel Panel, TextBox TextBox) LabeledTextBox(string label, string defaultText, bool multiline = false, int height = 30)
        {
            var panel = new Panel { Width = 940, Height = height };
            var lbl = new Label { Text = label, Left = 0, Top = 4, Width = 200 };
            var tb = new TextBox
            {
                Left = 210,
                Top = 0,
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
    }
}