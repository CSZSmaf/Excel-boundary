# 实现说明

Excel-boundary 由 C# COM 加载项和 C++ XLL 组成。COM 加载项响应工作簿事件，读取内容边界；XLL 在支持的 Office 构建中更新原生滚动和绘制边界。

## 内容边界

第一行为表头。内容扫描读取单元格显示值，配置中的 `IgnoredHeaders` 列不参与最后一行判断。这样可以保留预填序号和计数公式，同时让底部边界跟随实际记录。

末尾始终包含一个输入行。输入行写入内容后，扩展范围并在需要时调整滚动位置，让下一行靠近视口中部。边界控制不隐藏后续行列，也不删除或写入单元格。

列宽和行高以点数缓存。转换到显示尺寸时，按 DPI 和缩放逐行逐列取整，再求和，匹配 Excel 的显示边界。相同尺寸的连续行列作为一组缓存；总尺寸相同但组内尺寸发生变化时，仍会重新计算。

## 原生滚动

支持的 Office 版本通过 AirSpace / Windows Composition 的 `InteractionTracker` 驱动平滑滚动。XLL 修改相关表达式中的显示位移、最小位置和最大位置，并保留原有的惯性及 `IsPositionRoundingSuggested` 取整分支。

- 显示位移限制在 tracker 的最小、最大位置之间。
- 最小位置包含冻结行列的尺寸，防止冻结区域再次出现在可滚动区域。
- 水平方向最大位置由内容宽度减去视口宽度得到。
- 垂直方向最大位置由内容高度减去半个视口高度得到，使输入行靠近中部。
- 内容视觉对象的裁剪区域限制实际绘制范围，避免额外的空白行穿过边界。

表达式包含启用标志。切离配置的工作簿时，加载项清除边界尺寸，恢复原始画布尺寸和裁剪，并关闭表达式中的边界分支。

## 调度与资源

工作簿打开、激活和切换工作表事件触发目标选择。对非目标工作簿停止计时器，由后续激活事件重新唤醒。

目标工作簿位于前台时，轻量计时器观察原生位置变化。检测到滚动后推迟维护；输入结束后，最多每秒读取一次末尾两行，兼容临时关闭 Excel 事件的粘贴宏。尺寸检查有输入空闲条件和五秒间隔。

缓存的 Range 和 Window 引用在切换、停用时释放。原生滚动层销毁时释放对应的合成对象绑定，包括非工作表界面中的滚动层。调度使用单调时钟，支持长时间运行的 Excel 会话。

## 接口与构建

生产 COM 注册身份保留为 `Codex.ReadingSmooth`，用于兼容早期本地安装。显示名称为 **Excel Boundary**。测试构建使用独立的 `ExcelBoundary.Lab` 注册身份。

原生边界通过 Excel 进程内的消息窗口接收尺寸数组。生产版本不包含合成手势探针；探针和性能计数器仅编译进测试版本。

相关接口文档：[Excel COM 加载项](https://support.microsoft.com/en-us/excel/excel-com-add-ins-and-automation-add-ins)、[InteractionTracker](https://learn.microsoft.com/en-us/uwp/api/windows.ui.composition.interactions.interactiontracker)、[ExpressionAnimation](https://learn.microsoft.com/en-us/uwp/api/windows.ui.composition.expressionanimation)。
