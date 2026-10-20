using System.Globalization;
using AirDefenseLoad.Models;

namespace AirDefenseLoad.Services;

/// <summary>
/// 把计算结果铺成两张表（对应原程序 dataGridView2 与 dataGridView3），
/// 两者首尾相接就是 dataGridView4，也就是导出 Excel 的内容。
///
/// 单元格文案逐字照抄原程序 button3_Click 里的赋值，包括
/// 「K∑p= 0.9」这种带空格但没补零的写法 —— 那是 decimal/double 直接拼字符串的结果。
/// </summary>
public static class ResultTableBuilder
{
    public static string F2(double v) => v.ToString("0.00", CultureInfo.CurrentCulture);

    public static string F3(double v) => v.ToString("0.000", CultureInfo.CurrentCulture);

    /// <summary>明细表：每个单元两行（一级 / 二级），单元名只写在「一级」行。</summary>
    public static List<ResultRow> BuildDetail(CalculationResult r)
    {
        var rows = new List<ResultRow>(r.Units.Count * 2);

        foreach (var u in r.Units)
        {
            rows.Add(new ResultRow
            {
                Unit = u.Name,
                Level = "一级",
                Pn = F2(u.P1),
                Pc = F2(u.P1c),
                Qc = F2(u.Q1c),
                S = F2(u.S1),
                Remark = "k=" + u.K.ToString(CultureInfo.CurrentCulture)
            });

            rows.Add(new ResultRow
            {
                Unit = "",
                Level = "二级",
                Pn = F2(u.P2),
                Pc = F2(u.P2c),
                Qc = F2(u.Q2c),
                S = F2(u.S2),
                Remark = "cos=" + u.Cos.ToString(CultureInfo.CurrentCulture)
            });
        }

        return rows;
    }

    /// <summary>合计 / 同时系数 / 总计 / 综合系数，固定 6 行。</summary>
    public static List<ResultRow> BuildSummary(CalculationResult r)
    {
        return new List<ResultRow>
        {
            new()
            {
                Unit = "合计", Level = "一级",
                Pn = F2(r.SumP1), Pc = F2(r.SumP1c), Qc = F2(r.SumQ1c), S = F2(r.SumS1),
                Remark = ""
            },
            new()
            {
                Unit = "", Level = "二级",
                Pn = F2(r.SumP2), Pc = F2(r.SumP2c), Qc = F2(r.SumQ2c), S = F2(r.SumS2),
                Remark = ""
            },
            new()
            {
                Unit = "", Level = "一二级",
                Pn = F2(r.Pn), Pc = F2(r.SumPc), Qc = F2(r.SumQc), S = F2(r.SumS),
                Remark = ""
            },
            new()
            {
                Unit = "",
                Level = "K∑p= " + r.Kp.ToString(CultureInfo.CurrentCulture),
                Pn = "K∑q= " + r.Kq.ToString(CultureInfo.CurrentCulture),
                Remark = ""
            },
            new()
            {
                Unit = "总计",
                Level = "Pn= " + F2(r.Pn) + " kW",
                Pn = "Pc= " + F2(r.PcTotal) + " kW",
                Pc = "Sc= " + F2(r.Sc) + " kVA",
                Remark = ""
            },
            new()
            {
                Unit = "",
                Level = "K∑= " + F3(r.Kz),
                Pn = "cos∑= " + F3(r.CosZ),
                Pc = "Ic= " + F2(r.Ic) + " A",
                Remark = ""
            }
        };
    }
}
