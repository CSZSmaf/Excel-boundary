# 兼容性

当前适配版本：Microsoft Excel **16.0.20527.20196，x64**。

| 模块 | PE 时间戳 | 映像大小 |
| --- | --- | --- |
| EXCEL.EXE | `0x6af278b3` | `0x4ca7000` |
| mso40uiwin32client.dll | `0x6af2712e` | `0x138a000` |

加载时验证这些构建标识，再修改导入表和对应虚函数入口。构建不匹配时返回加载失败，Excel 保持普通滚动。

Office 更新后，需要重新确认滚动表达式、对象布局和入口位置，并运行原生边界与工作簿集成测试。适配信息集中在 `src/native/BoundaryExpressions.cpp`、`DynamicCanvas.h` 和 `FrozenPaneBoundary.h`。

当前工作流针对单个配置文件、单个活动工作表视图。冻结窗格、缩放和隐藏行列参与尺寸计算。多个工作簿窗口并排显示时，以当前活动工作簿作为启用范围。

诊断日志可在安装时通过 `-Diagnostics` 开启，也可将设置中的 `Diagnostics` 改为 `true` 后重新启动 Excel。日志位于 `%LOCALAPPDATA%\CodexExcelBoundary\logs`；正常运行默认关闭。
