using System.Globalization;
using AirDefenseLoad.Models;

namespace AirDefenseLoad.Services;

/// <summary>
/// 负荷计算内核 —— 从原 WinForm 的 button3_Click 中抽离出来的纯计算逻辑，
/// 表达式与运算顺序与原程序完全一致。
///
///   Pc = P × k
///   Qc = Pc × tan(arccos(cosφ))
///   S  = Pc / cosφ
///   Sc = √( (ΣPc·K∑p)² + (ΣQc·K∑q)² )
///   cos∑ = (ΣPc·K∑p) / Sc
///   K∑   = (ΣPc·K∑p) / ΣPn
///   Ic   = Sc / (1.732 × 0.38)
/// </summary>
public static class LoadCalculator
{
    private static readonly NumberStyles Styles = NumberStyles.Float | NumberStyles.AllowThousands;

    public static CalculationResult Calculate(
        IReadOnlyList<UnitRow> rows, double kp, double kq, out bool hasInvalidInput)
    {
        var details = new List<UnitCalculation>(rows.Count);
        hasInvalidInput = false;

        foreach (var row in rows)
        {
            double p1 = 0, p2 = 0, k = 0, cos = 0;

            // 原程序是顺序 Convert.ToDouble，任何一个转换失败就跳出该行的赋值，
            // 后面的值保持数组清零后的 0 —— 这里逐字照做。
            if (!TryNum(row.P1, out p1))
            {
                hasInvalidInput = true;
            }
            else if (!TryNum(row.P2, out p2))
            {
                hasInvalidInput = true;
            }
            else if (!TryNum(row.K, out k))
            {
                hasInvalidInput = true;
            }
            else if (!TryNum(row.Cos, out cos))
            {
                hasInvalidInput = true;
            }
            else
            {
                var tan = Math.Tan(Math.Acos(cos));
                var p1c = p1 * k;
                var p2c = p2 * k;
                details.Add(new UnitCalculation
                {
                    Name = row.Name,
                    P1 = p1,
                    P2 = p2,
                    K = k,
                    Cos = cos,
                    P1c = p1c,
                    P2c = p2c,
                    Q1c = p1c * tan,
                    Q2c = p2c * tan,
                    S1 = p1c / cos,
                    S2 = p2c / cos
                });
                continue;
            }

            // 该行有非法输入：原程序里这一行的 p1c/p2c/q1c/q2c 都停留在 0
            details.Add(new UnitCalculation
            {
                Name = row.Name,
                P1 = p1,
                P2 = p2,
                K = k,
                Cos = cos,
                P1c = 0,
                P2c = 0,
                Q1c = 0,
                Q2c = 0,
                S1 = 0,
                S2 = 0
            });
        }

        return new CalculationResult
        {
            Units = details,
            Kp = kp,
            Kq = kq,
            SumP1 = details.Sum(d => d.P1),
            SumP2 = details.Sum(d => d.P2),
            SumP1c = details.Sum(d => d.P1c),
            SumP2c = details.Sum(d => d.P2c),
            SumQ1c = details.Sum(d => d.Q1c),
            SumQ2c = details.Sum(d => d.Q2c),
            SumS1 = details.Sum(d => d.S1),
            SumS2 = details.Sum(d => d.S2)
        };
    }

    /// <summary>
    /// 模拟 Convert.ToDouble(string) 的行为：
    /// null / 空串按 0 处理（不报错），其余无法解析的文本视为非法输入。
    /// </summary>
    private static bool TryNum(string? text, out double value)
    {
        value = 0;
        if (string.IsNullOrWhiteSpace(text)) return true;
        return double.TryParse(text.Trim(), Styles, CultureInfo.CurrentCulture, out value);
    }
}
