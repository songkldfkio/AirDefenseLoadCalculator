using AirDefenseLoad.Models;

namespace AirDefenseLoad.Services;

/// <summary>
/// 导出负荷计算表到 .xlsx（内置 XlsxWriter 直接写 SpreadsheetML，无需安装 Office）。
///
/// 版式：标题行 + 浅蓝表头 + 细线表格；明细每单元两行（单元名纵向合并），
/// 底部接合计 / 总计共 6 行，整块浅蓝底提示。
/// 单元格内容与位置 = 界面 dataGridView4 的原样，不做任何重排。
/// </summary>
public static class ExcelExporter
{
    private const int ColCount = 7;

    /// <summary>表头文案（数值列主动折成两行，表头看起来更整齐）。</summary>
    public static readonly string[] Headers =
    {
        "单元",
        "负荷等级",
        "安装容量\nPn(kW)",
        "计算负荷\nPc(kW)",
        "计算负荷\nQc(kVar)",
        "视在负荷\nS(kVA)",
        "备注"
    };

    // ---------- 配色（ARGB） ----------
    private const string OuterLine = "FF8AA3BE";
    private const string InnerLine = "FFC9D6E4";
    private const string HeaderFill = "FFF1F5F9";
    private const string HeaderText = "FF33475B";
    private const string SumFill = "FFEDF4FD";

    private const string BodyFont = "宋体";
    private const string TitleFont = "微软雅黑";
    private const double BodySize = 10;

    /// <summary>
    /// 导出。
    /// </summary>
    /// <param name="path">目标 .xlsx 路径。</param>
    /// <param name="detail">明细行：每个防护单元两行（一级 / 二级）。</param>
    /// <param name="summary">底部汇总行：合计 3 行 + 总计 3 行。</param>
    public static void Export(string path, IReadOnlyList<ResultRow> detail, IReadOnlyList<ResultRow> summary)
    {
        var sheet = new XlsxWriter("负荷计算");

        const int titleRow = 1;
        const int headerRow = 2;

        BuildTitle(sheet, titleRow);
        BuildHeader(sheet, headerRow);

        // ---------- 明细 ----------
        var row = headerRow + 1;
        var detailStart = row;
        foreach (var r in detail) WriteDataRow(sheet, row++, r);

        // ---------- 合计 / 总计 ----------
        var sumStart = row;
        foreach (var r in summary) WriteDataRow(sheet, row++, r, isSummary: true);

        var lastRow = row - 1;

        // 明细的单元名跨两行居中；合计 / 总计各自留在自己那一行，
        // 与界面上的 dgv4 完全一致（不把「总计」拉到 K∑p 行上去）
        MergeUnitColumn(sheet, detailStart, detail);

        // ---------- 边框 ----------
        sheet.ApplyBorders(headerRow, 1, lastRow, ColCount, InnerLine, OuterLine);

        // 汇总区上方加一条稍重的分隔线
        if (summary.Count > 0)
        {
            sheet.SetTopBorder(sumStart, 1, ColCount, OuterLine, LineStyle.Medium);
        }

        // ---------- 列宽 / 冻结 ----------
        // 表头文案较长（「安装容量 Pn(kW)」），列宽留一点余量，避免字贴着边框
        var widths = new[] { 14.0, 16.0, 17.0, 17.0, 18.0, 17.0, 12.0 };
        for (var c = 0; c < ColCount; c++) sheet.SetColumnWidth(c + 1, widths[c]);

        sheet.FreezeRows(headerRow);

        sheet.Save(path);
    }

    // ================= 各部分 =================

    private static void BuildTitle(XlsxWriter sheet, int titleRow)
    {
        var style = new CellStyle
        {
            FontName = TitleFont,
            FontSize = 16,
            Bold = true,
            HAlign = HAlign.Center,
            VAlign = VAlign.Center
        };

        sheet.SetText(titleRow, 1, "防空地下室战时负荷计算表", style);
        sheet.Merge(titleRow, 1, titleRow, ColCount);
        sheet.SetRowHeight(titleRow, 34);
    }

    private static void BuildHeader(XlsxWriter sheet, int headerRow)
    {
        for (var c = 0; c < ColCount; c++)
        {
            var style = new CellStyle
            {
                FontName = BodyFont,
                FontSize = BodySize,
                Bold = true,
                FontColor = HeaderText,
                Fill = HeaderFill,
                HAlign = HAlign.Center,
                VAlign = VAlign.Center,
                Wrap = true
            };

            sheet.SetText(headerRow, c + 1, Headers[c], style);
        }

        sheet.SetRowHeight(headerRow, 34);
    }

    /// <summary>写一行数据：1/2/7 列居中的文本，3~6 列是数字就按数值写（可再参与 Excel 计算）。</summary>
    private static void WriteDataRow(XlsxWriter sheet, int row, ResultRow r, bool isSummary = false)
    {
        var cells = new[]
        {
            (Column: 1, Text: r.Unit),
            (Column: 2, Text: r.Level),
            (Column: 3, Text: r.Pn),
            (Column: 4, Text: r.Pc),
            (Column: 5, Text: r.Qc),
            (Column: 6, Text: r.S),
            (Column: 7, Text: r.Remark)
        };

        foreach (var (column, text) in cells)
        {
            var style = new CellStyle
            {
                FontName = BodyFont,
                FontSize = BodySize,
                VAlign = VAlign.Center,
                Bold = isSummary,
                Fill = isSummary ? SumFill : null
            };

            if (column is 3 or 4 or 5 or 6 && TryParseNumber(text, out var number))
            {
                style.NumberFormat = "0.00";
                style.HAlign = HAlign.Right;
                style.Indent = 1;
                sheet.SetNumber(row, column, number, style);
            }
            else
            {
                WriteTextCell(sheet, row, column, text, style);
            }
        }

        sheet.SetRowHeight(row, 20);
    }

    private static void WriteTextCell(XlsxWriter sheet, int row, int column, string? value, CellStyle style)
    {
        style.NumberFormat = "@";
        style.HAlign = HAlign.Center;
        sheet.SetText(row, column, value ?? string.Empty, style);
    }

    /// <summary>明细的「单元」列：每个防护单元占两行，名称纵向合并居中。</summary>
    private static void MergeUnitColumn(XlsxWriter sheet, int startRow, IReadOnlyList<ResultRow> detail)
    {
        for (var i = 0; i + 1 < detail.Count; i += 2)
        {
            var name = detail[i].Unit;
            if (string.IsNullOrEmpty(name)) name = detail[i + 1].Unit;
            if (string.IsNullOrEmpty(name)) continue;

            var style = new CellStyle
            {
                FontName = BodyFont,
                FontSize = BodySize,
                VAlign = VAlign.Center,
                HAlign = HAlign.Center,
                NumberFormat = "@"
            };

            sheet.SetText(startRow + i, 1, name, style);
            sheet.Merge(startRow + i, 1, startRow + i + 1, 1);
        }
    }

    private static bool TryParseNumber(string? text, out double value)
    {
        value = 0;
        if (string.IsNullOrWhiteSpace(text)) return false;

        return double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value)
               && !double.IsNaN(value) && !double.IsInfinity(value);
    }
}
