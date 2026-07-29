# Design Decisions

## 2025-10-28

**Decision**: 使用单文件 WinForms 架构

**Reason**: 项目规模适中，单文件便于快速开发和维护

**Impact**: 所有业务逻辑集中在 Form1.cs，Designer 管理 UI

**Status**: Active

---

**Decision**: 选择 .NET 8.0 + Windows Forms

**Reason**: 目标平台为 Windows，WinForms 开发效率高，适合工具类应用

**Impact**: 仅支持 Windows 平台，无法跨平台

**Status**: Active

---

**Decision**: 使用 async/await 处理所有网络操作

**Reason**: 避免 UI 阻塞，提升用户体验

**Impact**: 所有网络方法返回 Task，需要正确的异常处理

**Status**: Active

## 2026-07-28

**Decision**: Tab 布局选择 ListBox 左侧导航 + 隐藏 TabControl

**Reason**: 用户不喜欢顶部水平标签需要滚动；左侧垂直 Tab 的 OwnerDraw 方案中英文显示不一致；ListBox 方案最简洁，标题完整显示，无滚动问题

**Impact**: UI 从单个 TabControl 改为 SplitContainer/Panel + ListBox + 隐藏 TabControl

**Status**: Active

---

**Decision**: 设计器兼容方案 — InitializeComponent 只放基本属性

**Reason**: WinForms 设计器无法解析 lambda、循环、复杂方法调用；InitializeUI() 在设计模式下执行会报错

**Impact**: Form1.Designer.cs 精简为 50 行；所有 UI 创建逻辑移到 Form1.cs 的 InitializeUI()；构造函数用 LicenseManager.Designtime 判断

**Status**: Active

---

**Decision**: 导航面板使用双 Panel 方案替代 SplitContainer

**Reason**: SplitContainer 的 SplitterDistance 在某些情况下不生效；双 Panel（Dock=Left + Dock=Fill）更可靠

**Impact**: 左侧 Panel 固定 600px 宽，右侧 Panel 自动填充

**Status**: Active

## 2026-07-29

**Decision**: 标题过长采用混合方案 — 缩短文本 + 多行显示

**Reason**: Label 宽度 200px 约显示 16 个汉字，超出被 TextBox 遮挡。大部分标题去掉括号说明即可缩短；`自定义 Headers` 和 `端口（Port，可留空使用默认）` 的括号内容有实际指导意义，保留并改为多行显示

**Impact**: LabeledTextBox 方法新增 labelMultiline 参数；6 个标题文本调整

**Status**: Active

---

**Decision**: 发布为自包含单文件 exe（win-x64）+ zip 分发

**Reason**: 用户希望拿到 exe 直接运行，不依赖 .NET 运行时，不需要额外下载

**Impact**: 发布产物 75MB exe + 74MB zip（含原生 DLL），通过 `dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true` 构建

**Status**: Active
