using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;

namespace AirDefenseLoad.Collections;

/// <summary>
/// 支持「一次性替换全部内容」的 ObservableCollection。
///
/// 普通的 Clear() + 若干次 Add() 会让绑定的 DataGrid 收到 N+1 次集合变更通知，
/// 每收到一次就要做一遍行/列的失效与布局；单元数量到几十上百时，
/// 快速连点「增加单元」就会把 UI 线程堆死。
/// 这里改成先静默改内部列表，最后只发一次 Reset 通知，刷新只发生一次。
/// </summary>
public sealed class RangeObservableCollection<T> : ObservableCollection<T>
{
    private bool _suppressNotification;

    public RangeObservableCollection() { }

    public RangeObservableCollection(IEnumerable<T> items) : base(items) { }

    /// <summary>用新内容整体替换旧内容，只触发一次 Reset。</summary>
    public void ReplaceAll(IEnumerable<T> items)
    {
        _suppressNotification = true;
        try
        {
            Items.Clear();
            foreach (var item in items) Items.Add(item);
        }
        finally
        {
            _suppressNotification = false;
        }

        OnPropertyChanged(new PropertyChangedEventArgs(nameof(Count)));
        OnPropertyChanged(new PropertyChangedEventArgs("Item[]"));
        OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
    }

    /// <summary>一次追加多个元素，只触发一次 Reset。</summary>
    public void AddRange(IEnumerable<T> items)
    {
        var added = false;
        _suppressNotification = true;
        try
        {
            foreach (var item in items)
            {
                Items.Add(item);
                added = true;
            }
        }
        finally
        {
            _suppressNotification = false;
        }

        if (!added) return;

        OnPropertyChanged(new PropertyChangedEventArgs(nameof(Count)));
        OnPropertyChanged(new PropertyChangedEventArgs("Item[]"));
        OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
    }

    protected override void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
    {
        if (!_suppressNotification) base.OnCollectionChanged(e);
    }
}
