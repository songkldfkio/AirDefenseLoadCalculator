using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using AirDefenseLoad.Models;
using AirDefenseLoad.Services;
using AirDefenseLoad.ViewModels;
using Microsoft.Win32;

namespace AirDefenseLoad;

public partial class MainWindow : Window
{
    private readonly MainViewModel _vm = new();

    /// <summary>进入导出视图前的客户区尺寸，返回时原样还原。</summary>
    private double _returnWidth = 412;
    private double _returnHeight = 483;

    /// <summary>窗口边框（标题栏 + 左右下边框）的厚度，用来把“客户区尺寸”换算成窗口尺寸。</summary>
    private double _chromeW = 16;
    private double _chromeH = 39;
    private bool _chromeMeasured;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _vm;
        _vm.InvalidInput += OnInvalidInput;

        // 原程序 ShowIcon = false（标题栏不画图标），这里给个全透明图标占位
        Icon = CreateTransparentIcon();

        // 三张结果表每次重算都会重建行，重建后 DataGridView 的当前单元格落回 (0,0)，
        // 显示成蓝底白字的那一格 —— 这里保持一致，界面对比才逐像素相同。
        _vm.Units.CollectionChanged += (_, _) => KeepFirstCellSelected(InputGrid);
        _vm.DetailRows.CollectionChanged += (_, _) => KeepFirstCellSelected(DetailGrid);
        _vm.SummaryRows.CollectionChanged += (_, _) => KeepFirstCellSelected(SummaryGrid);
        _vm.ExportRows.CollectionChanged += (_, _) => KeepFirstCellSelected(ExportGrid);

        Loaded += OnLoaded;
    }

    /// <summary>
    /// 首次显示时量出窗口边框厚度，把默认客户区精确对齐到原来的 412×483。
    /// 之后再切视图、拉伸窗口，都用客户区尺寸 + 边框厚度来换算。
    /// </summary>
    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (_chromeMeasured) return;
        _chromeMeasured = true;

        if (ActualWidth > RootGrid.ActualWidth) _chromeW = ActualWidth - RootGrid.ActualWidth;
        if (ActualHeight > RootGrid.ActualHeight) _chromeH = ActualHeight - RootGrid.ActualHeight;

        // 窗口不设最小尺寸：允许一直拖到只剩标题栏，内容超出部分由 RootGrid 裁掉
        SetClientSize(412, 483);
    }

    /// <summary>按客户区尺寸设置窗口大小（最大化时先还原）。</summary>
    private void SetClientSize(double clientWidth, double clientHeight)
    {
        if (WindowState == WindowState.Maximized) WindowState = WindowState.Normal;

        Width = clientWidth + _chromeW;
        Height = clientHeight + _chromeH;
    }

    /// <summary>
    /// 把表格的当前单元格放回第一格（对应 WinForms 里 Rows.Add 之后的默认状态）。
    /// 同一个表格在同一轮里只排一次队，避免连点时堆积成百上千个回调。
    /// </summary>
    private static readonly HashSet<DataGrid> PendingSelect = new();

    private static void KeepFirstCellSelected(DataGrid grid)
    {
        if (!PendingSelect.Add(grid)) return;

        grid.Dispatcher.BeginInvoke(new Action(() =>
        {
            PendingSelect.Remove(grid);

            if (grid.SelectedCells.Count > 0) return;
            SelectFirstCell(grid);
        }), DispatcherPriority.Background);
    }

    /// <summary>选中第一行第一列（重建表格 / 切到导出视图时的默认当前单元格）。</summary>
    private static void SelectFirstCell(DataGrid grid)
    {
        if (grid.Items.Count == 0 || grid.Columns.Count == 0) return;

        var cell = new DataGridCellInfo(grid.Items[0], grid.Columns[0]);
        grid.CurrentCell = cell;
        grid.SelectedCells.Clear();
        grid.SelectedCells.Add(cell);
    }

    private static ImageSource CreateTransparentIcon()
        => BitmapSource.Create(16, 16, 96, 96, PixelFormats.Bgra32, null, new byte[16 * 16 * 4], 64);

    /// <summary>输入了非数字：原程序的提示文案。</summary>
    private void OnInvalidInput()
    {
        // 等单元格编辑事务结束再弹窗，避免重入
        Dispatcher.BeginInvoke(new Action(() =>
            MessageBox.Show(this, "请检查一二级负荷、需要系数、功率因数，应输入纯数字类型！")));
    }

    /// <summary>改表格前先把正在编辑的单元格提交，避免增删行时编辑事务冲突。</summary>
    private void CommitInputEdit()
    {
        try
        {
            InputGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            InputGrid.CommitEdit(DataGridEditingUnit.Row, true);
        }
        catch
        {
            // 没有在编辑时忽略
        }
    }

    #region 左上方一排按钮

    private void AddUnit_Click(object sender, RoutedEventArgs e)
    {
        CommitInputEdit();
        if (!_vm.AddUnit())
            MessageBox.Show(this, "无法增加，仅支持100组数据！");
    }

    private void RemoveUnit_Click(object sender, RoutedEventArgs e)
    {
        CommitInputEdit();

        var index = GetSelectedUnitIndex();      // 没有选中行时返回 -1，按原程序删最后一行
        var next = _vm.RemoveUnit(index);

        if (next >= 0) SelectInputCell(next);
    }

    /// <summary>当前选中的输入行下标；没有选中行时返回 -1。</summary>
    private int GetSelectedUnitIndex()
    {
        if (InputGrid.CurrentCell.IsValid && InputGrid.CurrentCell.Item is UnitRow current)
        {
            var i = _vm.Units.IndexOf(current);
            if (i >= 0) return i;
        }

        foreach (var cell in InputGrid.SelectedCells)
        {
            if (cell.Item is not UnitRow row) continue;
            var i = _vm.Units.IndexOf(row);
            if (i >= 0) return i;
        }

        return -1;
    }

    /// <summary>删除之后把选中格落到相邻行，方便连续删。</summary>
    private void SelectInputCell(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= InputGrid.Items.Count || InputGrid.Columns.Count == 0) return;

        InputGrid.Dispatcher.BeginInvoke(new Action(() =>
        {
            var row = Math.Min(rowIndex, InputGrid.Items.Count - 1);
            if (row < 0) return;

            var cell = new DataGridCellInfo(InputGrid.Items[row], InputGrid.Columns[0]);
            InputGrid.CurrentCell = cell;
            InputGrid.SelectedCells.Clear();
            InputGrid.SelectedCells.Add(cell);
            InputGrid.ScrollIntoView(InputGrid.Items[row], InputGrid.Columns[0]);
        }), DispatcherPriority.Loaded);
    }

    private void Renumber_Click(object sender, RoutedEventArgs e)
    {
        CommitInputEdit();
        _vm.Renumber();
    }

    private void ResetAll_Click(object sender, RoutedEventArgs e)
    {
        CommitInputEdit();
        _vm.ResetAll();
    }

    /// <summary>
    /// 「&gt;&gt;」展开 / 「&lt;&lt;」收起（对应 button9）。
    /// 宽度够 1118 就显示出右边的明细、合计表，不够就藏起来，
    /// 所以手动把窗口拉宽也会自动展开，不会出现表格被裁掉一半的样子。
    /// </summary>
    private void Expand_Click(object sender, RoutedEventArgs e)
    {
        CommitInputEdit();

        var wide = RightPanel.Visibility == Visibility.Visible;

        var w = wide ? 412 : Math.Max(1118, RootGrid.ActualWidth);
        var h = Math.Max(483, RootGrid.ActualHeight);

        SetClientSize(w, h);
    }

    /// <summary>客户区宽度跨过阈值时自动显示/收起右侧两张表。</summary>
    private void Panel1_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        // 阈值比 1118 留 8px 余量：窗口按 DPI 缩放后，1118 会被算成 1117.6，
        // 用精确的 1118 判断就会出现「点了展开、右侧却还是空的」。
        var showRight = Math.Round(Panel1.ActualWidth) >= 1110;

        if (showRight != (RightPanel.Visibility == Visibility.Visible))
        {
            RightPanel.Visibility = showRight ? Visibility.Visible : Visibility.Collapsed;

            // 收起时左栏吃掉整个宽度，展开时左栏固定回 412
            LeftColumn.Width = showRight ? new GridLength(412) : new GridLength(1, GridUnitType.Star);
            RightColumn.Width = showRight ? new GridLength(1, GridUnitType.Star) : new GridLength(0);

            ExpandButton.Content = showRight ? "<<" : ">>";
        }

        SyncInputLastColumn(showRight ? 412 : e.NewSize.Width);
    }

    /// <summary>
    /// 左栏比默认的 412 宽时，让输入表最后一列吃掉多余宽度，
    /// 否则窗口拉宽后表格右边会留一大片空白（看着像内容丢了）。
    /// 回到 412 时恢复固定 80，保证默认尺寸下的列宽与原版逐像素一致。
    /// </summary>
    /// <param name="leftWidth">左栏这一帧的实际宽度。</param>
    private void SyncInputLastColumn(double leftWidth)
    {
        if (InputGrid.Columns.Count < 5) return;

        // 前四列固定 80+81+80+81=322；表格左右各让出 4px 边距，
        // 再减掉竖直滚动条的 17px，剩下的就是 cos 列能占的宽度。
        var stretched = leftWidth > 412.5;
        var target = stretched ? Math.Max(80, leftWidth - 8 - 17 - 322) : 80;

        var last = InputGrid.Columns[4];
        if (last.Width.IsAbsolute && Math.Abs(last.Width.Value - target) < 0.5) return;

        last.Width = new DataGridLength(target);
    }

    #endregion

    #region 导出 / 返回

    /// <summary>导出 Excel，然后切到图集样式表格视图（对应 button8）。</summary>
    private void Export_Click(object sender, RoutedEventArgs e)
    {
        CommitInputEdit();
        _vm.Recalculate(immediate: true);

        if (_vm.DetailRows.Count == 0)
        {
            MessageBox.Show(this, "当前无数据可导出!");
            return;
        }

        var dialog = new SaveFileDialog
        {
            Title = "保存Excel文件",
            Filter = "Excel文件(*.xlsx)|*.xlsx",
            FilterIndex = 1,
            AddExtension = true,
            DefaultExt = ".xlsx",
            // 自动带一个不重名的文件名，省得每次手打
            FileName = $"防空地下室战时负荷计算_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
        };

        if (dialog.ShowDialog(this) == true)
        {
            try
            {
                ExcelExporter.Export(dialog.FileName, _vm.DetailRows, _vm.SummaryRows);
            }
            catch
            {
                MessageBox.Show(this, "此文件已被打开，无法保存！");
            }
        }

        // 原程序无论导出成功、失败还是取消，都会把原界面藏起来并切到表格视图
        _returnWidth = RootGrid.ActualWidth;
        _returnHeight = RootGrid.ActualHeight;

        Panel1.Visibility = Visibility.Collapsed;
        ExportPanel.Visibility = Visibility.Visible;
        SetClientSize(Math.Max(706, RootGrid.ActualWidth), Math.Max(681, RootGrid.ActualHeight));

        // 切过来之后把当前单元格落到左上角，和原程序一致
        Dispatcher.BeginInvoke(new Action(() => SelectFirstCell(ExportGrid)), DispatcherPriority.Loaded);
    }

    /// <summary>返回计算界面（对应 button10）。</summary>
    private void Back_Click(object sender, RoutedEventArgs e)
    {
        ExportPanel.Visibility = Visibility.Collapsed;
        Panel1.Visibility = Visibility.Visible;

        // 窗口不设最小尺寸，切回计算界面时原样恢复进入导出前的尺寸
        SetClientSize(_returnWidth, _returnHeight);
    }

    #endregion

    #region 表格复制（ctrl+A 全选、ctrl+C 复制到剪贴板）

    /// <summary>
    /// 原 WinForms 的 DataGridView 默认支持 ctrl+A 全选、ctrl+C 复制成 Tab 分隔文本，
    /// 粘到 Excel / CAD 里就是一张表。WPF 的 DataGrid 不行：SelectionMode=Single 时
    /// ctrl+A 根本不可用，SelectionUnit=Cell 时连 SelectAll() 都会抛异常，
    /// 所以这里自己实现「全选单元格」和「复制成 Tab 分隔文本」。
    /// </summary>
    private void ResultGrid_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (sender is not DataGrid grid) return;
        if ((Keyboard.Modifiers & ModifierKeys.Control) == 0) return;

        switch (e.Key)
        {
            case Key.A:
                SelectAllCells(grid);
                e.Handled = true;
                break;

            case Key.C:
                if (grid.SelectedCells.Count == 0) SelectAllCells(grid);
                CopySelection(grid);
                e.Handled = true;
                break;
        }
    }

    /// <summary>把整张表的单元格都标成选中（对应 WinForms 里 ctrl+A 的效果）。</summary>
    private static void SelectAllCells(DataGrid grid)
    {
        if (grid.Items.Count == 0 || grid.Columns.Count == 0) return;

        grid.SelectedCells.Clear();

        var all = new List<DataGridCellInfo>(grid.Items.Count * grid.Columns.Count);
        foreach (var item in grid.Items)
            foreach (var column in grid.Columns)
                all.Add(new DataGridCellInfo(item, column));

        foreach (var cell in all)
            grid.SelectedCells.Add(cell);
    }

    /// <summary>
    /// 把选中的单元格按「表头 + Tab 分隔」写进剪贴板，
    /// Excel 直接按列拆开，CAD 里粘贴则是一段带制表符的文本（与原程序一致）。
    /// </summary>
    private static void CopySelection(DataGrid grid)
    {
        if (grid.Items.Count == 0 || grid.Columns.Count == 0) return;

        int minRow = int.MaxValue, maxRow = -1, minCol = int.MaxValue, maxCol = -1;

        foreach (var cell in grid.SelectedCells)
        {
            var r = grid.Items.IndexOf(cell.Item);
            var c = grid.Columns.IndexOf(cell.Column);
            if (r < 0 || c < 0) continue;

            if (r < minRow) minRow = r;
            if (r > maxRow) maxRow = r;
            if (c < minCol) minCol = c;
            if (c > maxCol) maxCol = c;
        }

        // 没有选中任何单元格时按整张表复制
        if (maxRow < 0)
        {
            minRow = 0;
            maxRow = grid.Items.Count - 1;
            minCol = 0;
            maxCol = grid.Columns.Count - 1;
        }

        var sb = new StringBuilder();

        for (var c = minCol; c <= maxCol; c++)
        {
            if (c > minCol) sb.Append('\t');
            sb.Append(grid.Columns[c].Header);
        }
        sb.Append("\r\n");

        for (var r = minRow; r <= maxRow; r++)
        {
            var item = grid.Items[r];

            for (var c = minCol; c <= maxCol; c++)
            {
                if (c > minCol) sb.Append('\t');
                sb.Append(GetCellText(item, grid.Columns[c]));
            }
            sb.Append("\r\n");
        }

        try
        {
            Clipboard.SetText(sb.ToString());
        }
        catch
        {
            // 剪贴板被别的程序占用时忽略，不弹窗打断操作
        }
    }

    /// <summary>取单元格显示文本（数据源里的属性值，未实例化的虚拟行也能取到）。</summary>
    private static string GetCellText(object? item, DataGridColumn column)
    {
        if (item is null) return "";
        if (column is not DataGridBoundColumn bound) return "";
        if (bound.Binding is not Binding binding) return "";

        var path = binding.Path?.Path;
        if (string.IsNullOrEmpty(path)) return "";

        var prop = item.GetType().GetProperty(path);
        return prop?.GetValue(item)?.ToString() ?? "";
    }

    #endregion
}
