namespace AirDefenseLoad.Models;

/// <summary>
/// 结果表的一行（对应原程序 dataGridView2 / dataGridView3 / dataGridView4 的一行）。
/// 七个单元格全部是已经格式化好的字符串，直接绑定显示或写进 Excel。
/// </summary>
public sealed class ResultRow
{
    public string Unit { get; init; } = "";
    public string Level { get; init; } = "";
    public string Pn { get; init; } = "";
    public string Pc { get; init; } = "";
    public string Qc { get; init; } = "";
    public string S { get; init; } = "";
    public string Remark { get; init; } = "";
}
