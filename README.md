# 防空地下室战时负荷计算 V2.0

人防工程（防空地下室）战时负荷计算工具。输入各防护单元的安装容量、需要系数、功率因数，一键算出计算负荷、视在负荷与总电流，并导出带格式的 Excel 表。

本项目包含**两个版本**：原始的 WinForms 版（V1.1，作为参照实现保留）和用 WPF 重写的新版（V2.0）。

## 功能

- **多防护单元**：任意增删防护单元，删除后不重排序号，避免与图纸编号错位
- **三张结果表**：单元负荷明细、合计、总计
- **参数可调**：同时系数 K∑p / K∑q 可修改后重算
- **导出 Excel**：大标题、浅蓝表头、单元名纵向合并、冻结窗格、按一页宽打印缩放，文件名自动带时间戳（`防空地下室战时负荷计算_20261020_1135.xlsx`）
- **复制到剪贴板**：`ctrl+A` 全选、`ctrl+C` 复制，可直接粘进 Excel 或 CAD
- **窗口可自由缩放**：三个界面（计算 / 展开 / 导出）都能从最大化拖到只剩标题栏

## 计算式

```
Pc  = P × k
Qc  = Pc × tan(arccos(cosφ))
S   = Pc / cosφ
Sc  = √( (ΣPc·K∑p)² + (ΣQc·K∑q)² )
cos∑ = (ΣPc·K∑p) / Sc
K∑   = (ΣPc·K∑p) / ΣPn
Ic   = Sc / (1.732 × 0.38)
```

## 体积：78 KB

原版 exe 只有 42 KB。WPF 重写版最初用 .NET 8 自包含发布，膨胀到 74 MB。后来做了两件事把它压回 78 KB：

1. **目标框架改回 .NET Framework 4.8** —— Win10/11 系统自带，不需要打包运行时，用户零安装；
2. **去掉 ClosedXML** —— 那个全家桶独占 8.8 MB，改为自己实现的 `XlsxWriter`（约 480 行），直接按 SpreadsheetML 格式拼 XML 再塞进 ZipArchive，支持字体、底纹、边框、列宽行高、合并单元格、冻结窗格、横向打印缩放。

## 目录结构

```
.
├── Form1.cs / Form2.cs …      # 原 WinForms 版（V1.1），作为计算逻辑与布局的参照
├── 表格自动窗体.csproj
├── WpfApp/                    # WPF 重写版（V2.0），主开发目标
│   ├── MainWindow.xaml        # 主界面，Canvas 绝对定位复刻原版坐标
│   ├── Themes/Classic.xaml    # 把 DataGrid 重模板化成 DataGridView 外观
│   ├── ViewModels/            # MVVM，含 25ms 节流的合并重算
│   ├── Services/
│   │   ├── LoadCalculator.cs  # 计算内核（纯函数，与原版逐位一致）
│   │   ├── ResultTableBuilder.cs
│   │   ├── ExcelExporter.cs   # 版式定义
│   │   └── XlsxWriter.cs      # 自实现的 xlsx 写入器
│   └── dist/                  # 发布产物（构建生成，未入库）
└── .gitignore
```

## 构建与运行

需要 Windows + .NET Framework 4.8（系统自带）+ .NET SDK 6 以上（用于编译）。

```powershell
cd WpfApp
dotnet build -c Release
dotnet publish -c Release -o dist
```

产物在 `WpfApp/dist/防空地下室战时负荷计算.exe`，单文件、免安装，直接双击运行。

原 WinForms 版的构建需要用 Visual Studio 打开 `MID多窗体.sln`，第一次还原会拉取 `packages.config` 里的 NuGet 包。

## 验证方式

重写过程中以原版 exe 为基准做了自动化对照：

- **数值**：用反射加载原版 exe，驱动按钮读取 `dataGridView4` 文本，1907 组随机与边界样例，**0 处差异**
- **界面**：四种状态（默认 412×483 / 展开 1118×483 / 导出 706×681 / 各尺寸拉伸）截图逐像素比对，剩余差异仅为 GDI 与 WPF 的文字抗锯齿
- **性能**：连点 100 次 6.8s、全部重置 468ms、连删 100 次 258ms；100 单元全选复制 207 行 0.15s
- **导出**：生成的 xlsx 用 WPS 打开目检，并用脚本回读单元格确认排版与数值

## 已知限制

- 仅支持 Windows
- 首次点「导出」时保存对话框要等 1～3 秒（系统首次加载对话框组件），之后就秒开
