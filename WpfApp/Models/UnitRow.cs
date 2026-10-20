namespace AirDefenseLoad.Models;

/// <summary>
/// 一个防护单元的输入行（对应原程序 dataGridView1 的一行）。
///
/// 四个数值列一律按「文本」保存：原程序就是往单元格里写字符串，
/// 界面显示的是用户敲进去的原样（"7.5" 而不是 "7.50"），
/// 参与计算的则是 <see cref="double.Parse"/> 之后的数值。
/// </summary>
public sealed class UnitRow : ObservableObject
{
    public const string DefaultLevel1 = "7.5";
    public const string DefaultLevel2 = "10.5";
    public const string DefaultDemandFactor = "0.9";
    public const string DefaultPowerFactor = "0.8";

    private string _name = "";
    private string _p1 = DefaultLevel1;
    private string _p2 = DefaultLevel2;
    private string _k = DefaultDemandFactor;
    private string _cos = DefaultPowerFactor;

    public UnitRow() { }

    public UnitRow(string name) => _name = name;

    /// <summary>单元名称，如「防护单元1」。</summary>
    public string Name
    {
        get => _name;
        set => Set(ref _name, value ?? "");
    }

    /// <summary>一级负荷安装容量 (kW) 原文。</summary>
    public string P1
    {
        get => _p1;
        set => Set(ref _p1, value ?? "");
    }

    /// <summary>二级负荷安装容量 (kW) 原文。</summary>
    public string P2
    {
        get => _p2;
        set => Set(ref _p2, value ?? "");
    }

    /// <summary>需要系数 k 原文。</summary>
    public string K
    {
        get => _k;
        set => Set(ref _k, value ?? "");
    }

    /// <summary>功率因数 cosφ 原文。</summary>
    public string Cos
    {
        get => _cos;
        set => Set(ref _cos, value ?? "");
    }

    /// <summary>恢复默认值（保留单元名）。</summary>
    public void ResetToDefault()
    {
        P1 = DefaultLevel1;
        P2 = DefaultLevel2;
        K = DefaultDemandFactor;
        Cos = DefaultPowerFactor;
    }
}
