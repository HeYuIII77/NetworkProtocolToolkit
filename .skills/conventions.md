# Project Conventions

## 命名规范
- **私有字段**: `_camelCase` (如 `_logBox`, `_httpClient`, `_logDir`)
- **私有方法**: `PascalCase` (如 `EnsureLogDirectory`, `WriteToDailyLog`)
- **异步方法**: `Do` + 动词 或 `Test` + 名词 (如 `DoHttpGet`, `TestIpPort`)
- **UI 创建方法**: `Create` + 功能名 + `Tab` (如 `CreateNetworkDiagTab`)

## 代码风格
- 使用 `var` 进行类型推断
- XML 文档注释覆盖所有方法
- 异常处理：try-catch 包裹网络操作，静默失败或显示 MessageBox
- 日志分级：INFO / WARN / ERROR

## 文件组织
- `Form1.Designer.cs`: 只放 InitializeComponent 基本属性（设计器兼容）
- `Form1.cs`: 所有 UI 创建（InitializeUI + Create*Tab）和业务逻辑
- `Program.cs`: 入口点

## 提交规范
- 中文提交信息
- 格式：`日期 + 描述` (如 `20251029 1、补充注释`)

## UI 规范
- 左侧 ListBox 导航（600px 宽）+ 右侧隐藏 TabControl 内容区
- LabeledTextBox 辅助方法创建一致的输入控件（Label 200px + TextBox 700px）
- 每个功能页有独立的响应显示区域
- Consolas 字体用于日志和响应显示
- 窗体固定大小（FixedSingle），不可最大化

## 设计器兼容规范
- InitializeComponent() 只放简单属性赋值
- 复杂 UI 创建放 InitializeUI()
- 构造函数用 `LicenseManager.UsageMode != LicenseUsageMode.Designtime` 判断
