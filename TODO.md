# TODO List

## v5.2.0 计划功能

### 新增功能

- [ ] **PPT转换选项增强**
  - 添加更多PPT转换相关的配置选项
  - 支持多种输出视图模式：
    - **讲义模式**（一页显示多个幻灯片，如 2、3、4、6、9 张幻灯片/页）
    - **备注页模式**（幻灯片 + 演讲者备注）
    - **大纲视图模式**（仅显示文本大纲）
  - 可能还包括：
    - 幻灯片范围选择

- [ ] **右键菜单支持**
  - 在 Windows 资源管理器中添加右键菜单
  - 支持直接右键文件/文件夹快速转换
  - 集成到系统上下文菜单
  - 可能的菜单项：
    - "转换为PDF"
    - "批量转换为PDF"

### 技术实现要点

- PPT选项：扩展 `IOfficeApplication` 接口
- 右键菜单：需要注册表操作，注意权限问题
- 安装程序：可能需要创建安装程序来注册右键菜单

---

## 未来版本考虑

### 功能增强
- [ ] 支持更多文件格式（如 RTF、ODT 等）

### 性能优化
- [ ] 多线程并行转换

### 用户体验
- [ ] 添加转换进度条
- [ ] 多语言支持（英文界面）
- [ ] 拖拽文件直接转换

---

## 已知问题 / Bug修复

### 网络路径支持不完整

- [ ] **MSWordApplication（自动引擎-Word）缺少输出网络路径处理**
  - 问题：`SaveAsPDF` 方法只处理了输入网络路径，但没有处理输出到网络路径的情况
  - 影响：当输出目标是网络路径时（如 `\\server\share\output.pdf`），可能会因为网络延迟或权限问题导致转换失败
  - 解决方案：参考 MSExcelApplication 的实现，在 SaveAsPDF 中添加：
    ```csharp
    bool isNetworkOutput = NetworkPathHelper.IsNetworkPath(toFilePath);
    string actualOutputPath = isNetworkOutput ? NetworkPathHelper.CreateLocalTempOutputPath(toFilePath) : toFilePath;
    // ... 转换到本地临时路径 ...
    if (isNetworkOutput) {
        NetworkPathHelper.CopyToNetworkPath(actualOutputPath, toFilePath);
        NetworkPathHelper.CleanupTempFile(actualOutputPath);
    }
    ```

- [ ] **WpsWriterApplication（WPS文字引擎）完全没有网络路径处理**
  - 问题：无论输入还是输出都没有网络路径支持
  - 影响：
    - 输入网络路径时，WPS 可能无法直接打开文件或响应缓慢
    - 输出网络路径时，保存可能失败或耗时过长
  - 解决方案：参考 WpsSpreadsheetApplication 的完整实现
    - OpenDocument: 检查并创建本地临时副本
    - SaveAsPDF: 先输出到本地，再复制到网络路径
    - Dispose: 清理临时文件

---

_最后更新：2025-10-03_
