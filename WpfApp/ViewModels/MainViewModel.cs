using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Threading;
using AirDefenseLoad.Collections;
using AirDefenseLoad.Models;
using AirDefenseLoad.Services;

namespace AirDefenseLoad.ViewModels;

/// <summary>
/// 主界面数据。对应原程序里散落在 Form1 各处的那一堆数组和 grid 赋值，
/// 这里统一成「输入行集合 → 计算 → 三张结果表」的单向数据流。
///
/// 性能约定（这一版专门为「快速连点增加单元不卡」做的）：
///  1. 计算结果先在内存里算完，UI 集合用 ReplaceAll 一次性整块替换，只发一次变更通知；
///  2. 高频操作（连点、批量重置、序号重排）把刷新合并到空闲时执行，一帧只刷一次；
///  3. 批量改属性（ResetAll / Renumber）期间挂起重算，避免每改一个字段就全表重算。
/// </summary>
public sealed class MainViewModel : ObservableObject
{
    public const int MaxUnits = 100;

    /// <summary>两次 UI 刷新之间的最小间隔（毫秒）。连点时按这个频率合并。</summary>
    private const int FlushIntervalMs = 25;

    private static readonly List<ResultRow> EmptyRows = new(0);

    private readonly Dispatcher _dispatcher = Dispatcher.CurrentDispatcher;

    private double _kp = 0.9;
    private double _kq = 0.9;
    private CalculationResult _result = new();

    /// <summary>是否出现过单元。原程序启动时三张结果表都是空的，</summary>
    /// <summary>只有真正算过（哪怕后来把单元删到 0 个）才会留下那 6 行合计（0.00 / NaN）。</summary>
    private bool _hasHadUnits;

    // 待刷新的行缓存 + 合并调度状态
    private List<ResultRow> _pendingDetail = EmptyRows;
    private List<ResultRow> _pendingSummary = EmptyRows;
    private List<ResultRow> _pendingExport = EmptyRows;
    private int _version;
    private int _flushedVersion;
    private bool _flushScheduled;
    private long _lastFlushTick;

    /// <summary>批量修改属性期间挂起重算。</summary>
    private bool _suspendRecalc;

    public MainViewModel()
    {
        Units.CollectionChanged += OnUnitsChanged;
    }

    /// <summary>输入行（dataGridView1）。</summary>
    public ObservableCollection<UnitRow> Units { get; } = new();

    /// <summary>明细表（dataGridView2）：每个单元两行。</summary>
    public RangeObservableCollection<ResultRow> DetailRows { get; } = new();

    /// <summary>合计表（dataGridView3）：固定 6 行。</summary>
    public RangeObservableCollection<ResultRow> SummaryRows { get; } = new();

    /// <summary>导出表（dataGridView4）：明细 + 合计，导出 Excel 用的就是这张。</summary>
    public RangeObservableCollection<ResultRow> ExportRows { get; } = new();

    /// <summary>输入了非数字时触发（对应原程序的 MessageBox 提示）。</summary>
    public event Action? InvalidInput;

    /// <summary>有功同时系数 K∑p。</summary>
    public double Kp
    {
        get => _kp;
        set
        {
            if (!Set(ref _kp, value)) return;
            Recalculate();
        }
    }

    /// <summary>无功同时系数 K∑q。</summary>
    public double Kq
    {
        get => _kq;
        set
        {
            if (!Set(ref _kq, value)) return;
            Recalculate();
        }
    }

    public CalculationResult Result => _result;

    #region 按钮动作

    /// <summary>增加一个单元（对应 button1）。超过 100 组返回 false。</summary>
    public bool AddUnit()
    {
        if (Units.Count >= MaxUnits) return false;

        Units.Add(new UnitRow("防护单元" + NextUnitNumber()));
        return true;
    }

    /// <summary>
    /// 删除指定下标的单元（对应 button5）。
    /// 下标无效（没有选中行）时删除最后一个，和原程序一致。
    /// 删除后<strong>不</strong>重排「防护单元N」的序号，剩下的单元保持原编号。
    /// 返回删除后应当选中的行号（没有行时返回 -1）。
    /// </summary>
    public int RemoveUnit(int index)
    {
        if (Units.Count == 0) return -1;

        if (index < 0 || index >= Units.Count) index = Units.Count - 1;

        Units.RemoveAt(index);
        Recalculate();

        return Units.Count == 0 ? -1 : Math.Min(index, Units.Count - 1);
    }

    /// <summary>
    /// 新增单元用的编号：取现有「防护单元N」里最大的 N 再加 1。
    /// 原版直接用「行数 + 1」，但那只在「只能删最后一行」时成立；
    /// 现在支持删中间行且删后不重排，按行数 + 1 会撞出重复编号。
    /// 第一列用户可以自己改名，解析不出来的行忽略；一行都没有时退回 1。
    /// </summary>
    private int NextUnitNumber()
    {
        const string prefix = "防护单元";
        var max = 0;

        foreach (var unit in Units)
        {
            var name = unit.Name;
            if (string.IsNullOrEmpty(name) || !name.StartsWith(prefix, StringComparison.Ordinal)) continue;

            if (int.TryParse(name.Substring(prefix.Length), out var n) && n > max) max = n;
        }

        return max > 0 ? max + 1 : Units.Count + 1;
    }

    /// <summary>删除最后一个单元（原程序的行为，作为未选中行时的兜底）。</summary>
    public int RemoveLastUnit() => RemoveUnit(Units.Count - 1);

    /// <summary>按当前顺序重排单元名（对应 button7）。</summary>
    public void Renumber()
    {
        RenumberNames();
        Recalculate();
    }

    /// <summary>全部重置（对应 button2）：所有输入回默认值，同时系数回 0.90。</summary>
    public void ResetAll()
    {
        if (Units.Count == 0 && !_hasHadUnits) return;

        SuspendRecalc(() =>
        {
            foreach (var u in Units) u.ResetToDefault();
        });

        _kp = 0.9;
        _kq = 0.9;
        OnPropertyChanged(nameof(Kp));
        OnPropertyChanged(nameof(Kq));

        Recalculate();
    }

    #endregion

    #region 计算

    private void OnUnitsChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        if (e.OldItems is not null)
            foreach (UnitRow u in e.OldItems) u.PropertyChanged -= OnUnitPropertyChanged;

        if (e.NewItems is not null)
            foreach (UnitRow u in e.NewItems) u.PropertyChanged += OnUnitPropertyChanged;

        if (!_suspendRecalc) Recalculate();
    }

    private void OnUnitPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_suspendRecalc) return;
        Recalculate(fromCellEdit: true);
    }

    /// <summary>在挂起重算的情况下执行一批属性修改，结束后统一重算。</summary>
    private void SuspendRecalc(Action action)
    {
        var old = _suspendRecalc;
        _suspendRecalc = true;
        try
        {
            action();
        }
        finally
        {
            _suspendRecalc = old;
        }
    }

    private void RenumberNames()
    {
        SuspendRecalc(() =>
        {
            for (var i = 0; i < Units.Count; i++)
                Units[i].Name = "防护单元" + (i + 1);
        });
    }

    /// <summary>
    /// 重新计算并刷新三张表。
    /// 计算是同步完成的（<see cref="Result"/> 立刻可用），
    /// 但写入 UI 集合会按 25ms 合并，连续操作时只刷最后一帧。
    /// </summary>
    /// <param name="fromCellEdit">是不是单元格编辑触发的（只有这种才弹非法输入提示）。</param>
    /// <param name="immediate">true = 立刻写入集合（导出前必须用这个）。</param>
    public void Recalculate(bool fromCellEdit = false, bool immediate = false)
    {
        if (Units.Count == 0 && !_hasHadUnits)
        {
            _result = new CalculationResult();
            _pendingDetail = EmptyRows;
            _pendingSummary = EmptyRows;
            _pendingExport = EmptyRows;
            _version++;
        }
        else
        {
            if (Units.Count > 0) _hasHadUnits = true;
            if (_suspendRecalc) return;

            var result = LoadCalculator.Calculate(Units, Kp, Kq, out var hasInvalidInput);
            _result = result;

            var detail = ResultTableBuilder.BuildDetail(result);
            var summary = ResultTableBuilder.BuildSummary(result);

            _pendingDetail = detail;
            _pendingSummary = summary;

            var all = new List<ResultRow>(detail.Count + summary.Count);
            all.AddRange(detail);
            all.AddRange(summary);
            _pendingExport = all;

            _version++;

            if (hasInvalidInput && fromCellEdit) InvalidInput?.Invoke();
        }

        if (immediate) Flush();
        else ScheduleFlush();
    }

    /// <summary>把待刷新的行立刻写进三张表。</summary>
    public void Flush()
    {
        _flushScheduled = false;
        if (_flushedVersion == _version) return;

        _flushedVersion = _version;
        _lastFlushTick = NowMs();

        DetailRows.ReplaceAll(_pendingDetail);
        SummaryRows.ReplaceAll(_pendingSummary);
        ExportRows.ReplaceAll(_pendingExport);
    }

    /// <summary>毫秒级单调时钟。net48 没有 Environment.TickCount64，用 Stopwatch 顶替。</summary>
    private static long NowMs() => (long)(Stopwatch.GetTimestamp() * 1000.0 / Stopwatch.Frequency);

    /// <summary>合并调度：一批连续操作只刷一次界面，两次刷新至少间隔 25ms。</summary>
    private void ScheduleFlush()
    {
        if (_flushScheduled) return;
        _flushScheduled = true;
        _ = FlushCoalescedAsync();
    }

    private async Task FlushCoalescedAsync()
    {
        try
        {
            var since = NowMs() - _lastFlushTick;
            if (since >= 0 && since < FlushIntervalMs)
                await Task.Delay((int)(FlushIntervalMs - since)).ConfigureAwait(true);

            await _dispatcher.InvokeAsync(Flush, DispatcherPriority.Background);
        }
        catch (Exception)
        {
            // 窗口已经关闭 / Dispatcher 已停：忽略
            _flushScheduled = false;
        }
    }

    #endregion
}
