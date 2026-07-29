# Known Bugs

_(无已知 Bug)_

## 已解决

### OPC UA NullReferenceException
**Status**: Resolved — 2026-07-28 Designer 中 OPC UA 按钮事件只赋值了 _deviceProtoResponseBox，漏赋值 _opcUaRespBox，导致 DoOpcUaRead 访问 null。添加 `_opcUaRespBox = uaResp` 修复。

### 端口扫描卡 UI
**Status**: Resolved — 2026-07-28 扫描过程中无任何反馈。改为每 100 个端口更新一次 UI 显示进度，添加停止按钮和 CancellationTokenSource 取消机制。

### 设计器报错 IUIService
**Status**: Resolved — 2026-07-28 InitializeComponent 中包含复杂逻辑（lambda、循环），设计器无法解析。精简 InitializeComponent 为基本属性，复杂逻辑移到 InitializeUI()，设计模式下跳过。

### Oracle 数据库驱动缺失
**Status**: Resolved — 2026-07-28 通过反射动态检测，未安装时提示安装建议，不影响其他功能
