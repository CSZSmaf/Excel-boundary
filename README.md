# Excel-boundary

保留 Excel 的原生触控板平滑滚动，为工作表设置贴合内容的硬边界。

我们希望用触控板浏览和填写长表格时，滑动到边缘就自然停住，新增记录时也能在舒适的位置继续输入。Excel-boundary 将滚动范围限制在实际内容附近，处理边缘回弹带来的空白、冻结表头重复显示，以及输入行贴着窗口底部的问题。

## 功能

- **原生平滑滚动**：保留连续的像素位移、惯性和 Excel 自身的取整规则。
- **四边硬边界**：左侧和顶部停止越界；右侧对齐内容的最后一列。
- **底部自动延伸**：最后一条记录后露出一个输入行。填入新内容后，边界随记录延伸。
- **居中填写**：到达底部时，让输入行停在视口中部附近，留出继续输入的空间。
- **冻结窗格适配**：计算冻结行、列的边界，避免滑动时出现重复表头或重复首列。
- **普通方式打开**：安装后直接双击配置的 `.xlsx` 或 `.xlsm` 文件即可启用。
- **按工作簿启用**：切到其他工作簿时恢复普通边界并停止维护定时器。
- **编辑期间保持视图**：输入文字、修改公式或选取引用时保留已有边界，完成编辑后再更新。

插件读取工作表的内容和尺寸，边界随用户的编辑更新。现有工作簿中的日期、计数、粘贴整理等宏继续由工作簿自身管理。

**v0.1.1** 修复了编辑期间可能只显示 A1:A2、其余区域变灰的问题。请使用最新安装包更新，具体变更见 [更新记录](CHANGELOG.md)。

## 适用环境

当前版本适配 **Windows x64、64 位 Microsoft Excel 16.0.20527.20196**，使用 .NET Framework 4.8。需要桌面版 Excel。

原生部分对接这一版本的 Office 滚动实现，会核对 Excel 和 `mso40uiwin32client.dll` 的构建标识。其他构建会跳过原生挂接，按普通 Excel 运行。Office 更新后的适配情况见 [兼容性说明](docs/compatibility.md)。

## 安装

从 [Releases](https://github.com/CSZSmaf/Excel-boundary/releases) 下载 Windows x64 安装包并解压。保存并关闭 Excel，然后在解压目录打开 **64 位 PowerShell**：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 `
  -WorkbookPath "C:\Work\Reading.xlsx" -WorksheetName "Reading"
```

将路径和工作表名换成自己的文件。省略 `-WorksheetName` 时，对该工作簿当前选中的工作表启用。

如果表格包含预填的序号、计数或公式列，可以指定这些列的表头，让它们不参与“最后一条记录”的判断。例如：

```powershell
.\Install.ps1 -WorkbookPath "C:\Work\Reading.xlsm" `
  -WorksheetName "Reading" -IgnoredHeaders "Read count","Record ID"
```

表头按第一行读取，也可以使用中文表头名称。安装完成后，照常双击打开工作簿。

安装只使用当前用户的目录和加载项注册项。配置保存在 `%LOCALAPPDATA%\CodexExcelBoundary\ReadingAddin\settings.json`，示例见 [settings.example.json](settings.example.json)。更新文件路径或配置后，重新启动 Excel 即可生效。

## 停用与卸载

临时停用：Excel → 文件 → 选项 → 加载项 → 管理 **COM 加载项** → 转到，取消勾选 **Excel Boundary**。

卸载：保存并关闭 Excel，在安装包目录运行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Uninstall.ps1
```

## 运行开销

维护代码运行在 Excel 进程内。切换到未配置的工作簿后，维护定时器停止；Excel 位于后台时，跳过边界查询和表格扫描；滚动和惯性移动期间，暂停尺寸维护。平稳停下后按需检查输入行，尺寸检查最多每五秒执行一次。相同尺寸的行列复用缓存，常规运行默认关闭诊断日志。

这些策略用于减少边界维护的开销，同时保持 Excel 原生的滚动过程。验证方法和结果见 [测试说明](docs/testing.md)。

## 从源码构建

需要 Visual Studio Build Tools 的 **Desktop development with C++** 工作负载、包含 C++/WinRT 的 Windows SDK，以及 .NET Framework 的 C# 编译器。

```powershell
.\scripts\build.ps1
.\scripts\Install.ps1 -WorkbookPath "C:\Work\Reading.xlsx" -WorksheetName "Reading"
```

脚本自动查找 Visual Studio、Windows SDK 和 Excel Interop。未找到 Interop 时，会下载固定版本的 [NuGet 包](https://www.nuget.org/packages/microsoft.office.interop.excel/15.0.4795.1001)。Interop 类型嵌入生成的程序集，安装包只携带本项目的二进制文件。

实现说明见 [架构文档](docs/architecture.md)。采用 [MIT License](LICENSE)。
