using System.Globalization;

namespace AirDefenseLoad.Models;

/// <summary>单个防护单元的计算结果。</summary>
public sealed class UnitCalculation
{
    public string Name { get; init; } = "";
    public double P1 { get; init; }
    public double P2 { get; init; }
    public double K { get; init; }
    public double Cos { get; init; }

    /// <summary>一级计算有功 Pc1 = P1 × k (kW)</summary>
    public double P1c { get; init; }
    /// <summary>二级计算有功 Pc2 = P2 × k (kW)</summary>
    public double P2c { get; init; }
    /// <summary>一级计算无功 Qc1 = Pc1 × tanφ (kVar)</summary>
    public double Q1c { get; init; }
    /// <summary>二级计算无功 Qc2 = Pc2 × tanφ (kVar)</summary>
    public double Q2c { get; init; }
    /// <summary>一级视在功率 S1 = Pc1 / cosφ (kVA)</summary>
    public double S1 { get; init; }
    /// <summary>二级视在功率 S2 = Pc2 / cosφ (kVA)</summary>
    public double S2 { get; init; }
}

/// <summary>
/// 整张表的计算结果（对应原程序 button3_Click 算出来的全部数值）。
///
/// 注意这里所有表达式都照抄原程序，连运算顺序都不改：
/// 原程序是先把明细值按 "0.00" 取整写进单元格，再把这些单元格读回来求和，
/// 所以 Sc / cos∑ 用的是「按显示文本取整后」的明细合计；
/// 而 K∑ 的分子用的是「未取整」的 ΣPc，两者口径不同，不能统一。
/// </summary>
public sealed class CalculationResult
{
    public IReadOnlyList<UnitCalculation> Units { get; init; } = Array.Empty<UnitCalculation>();

    public double Kp { get; init; }
    public double Kq { get; init; }

    public double SumP1 { get; init; }
    public double SumP2 { get; init; }
    public double SumP1c { get; init; }
    public double SumP2c { get; init; }
    public double SumQ1c { get; init; }
    public double SumQ2c { get; init; }
    public double SumS1 { get; init; }
    public double SumS2 { get; init; }

    /// <summary>安装总容量 Pn = ΣP1 + ΣP2 (kW)</summary>
    public double Pn => SumP1 + SumP2;

    /// <summary>一二级计算有功合计 ΣPc（未取整）</summary>
    public double SumPc => SumP1c + SumP2c;

    /// <summary>一二级计算无功合计 ΣQc（未取整）</summary>
    public double SumQc => SumQ1c + SumQ2c;

    /// <summary>一二级视在功率算术和 ΣS</summary>
    public double SumS => SumS1 + SumS2;

    /// <summary>明细行 Pc 按 "0.00" 取整后的合计（= 原程序里变量 m）</summary>
    public double SumPcRounded => Units.Sum(u => LegacyRound2(u.P1c) + LegacyRound2(u.P2c));

    /// <summary>明细行 Qc 按 "0.00" 取整后的合计（= 原程序里变量 n）</summary>
    public double SumQcRounded => Units.Sum(u => LegacyRound2(u.Q1c) + LegacyRound2(u.Q2c));

    /// <summary>计入同时系数的计算有功 Pc = m × K∑p (kW)</summary>
    public double PcTotal => SumPcRounded * Kp;

    /// <summary>计入同时系数的计算无功 Qc = n × K∑q (kVar)</summary>
    public double QcTotal => SumQcRounded * Kq;

    /// <summary>计算视在功率 Sc = √(m²Kp² + n²Kq²) (kVA)</summary>
    public double Sc => Math.Sqrt(SumPcRounded * SumPcRounded * Kp * Kp
                                + SumQcRounded * SumQcRounded * Kq * Kq);

    /// <summary>
    /// 综合需要系数 K∑ = (ΣP1c·K∑p + ΣP2c·K∑p) / Pn（分子不取整）。
    /// 不加 0 保护 —— 原程序在 Pn=0 时除出 NaN，单元格就显示 "NaN"。
    /// </summary>
    public double Kz => (SumP1c * Kp + SumP2c * Kp) / Pn;

    /// <summary>综合功率因数 cos∑ = (m × K∑p) / Sc（同样不做 0 保护，保留 NaN）</summary>
    public double CosZ => (SumPcRounded * Kp) / Sc;

    /// <summary>计算电流 Ic = Sc / (1.732 × 0.38) (A)</summary>
    public double Ic => Sc / (1.732 * 0.38);

    /// <summary>
    /// 旧版的「显示文本取整」：先按 "0.00" 格式化成文本，再解析回数值。
    /// 与 Math.Round 在极少数临界值上差 0.01（例：39.974999999999994 会格式化成 "39.98"）。
    /// 必须保持这个口径，总计才与旧版逐位一致。
    /// </summary>
    private static double LegacyRound2(double v) =>
        double.Parse(v.ToString("0.00", CultureInfo.CurrentCulture), CultureInfo.CurrentCulture);
}
