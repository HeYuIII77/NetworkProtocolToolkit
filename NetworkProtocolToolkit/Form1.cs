using System.Net;
using System.Net.Mail;
using System.Net.Sockets;
using System.Net.WebSockets;
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
            EnsureLogDirectory();
            AppendLog("应用启动", "INFO");
        }

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