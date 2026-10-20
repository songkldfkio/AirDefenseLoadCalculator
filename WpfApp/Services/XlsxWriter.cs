using System.IO.Compression;

namespace AirDefenseLoad.Services;

/// <summary>
/// 极简 .xlsx 写出器：直接按 SpreadsheetML 拼 XML，再压进 zip。
///
/// 只实现本程序用得到的部分（文本/数字、字体、底色、四边线、列宽行高、
/// 合并单元格、冻结窗格、横向打印按宽缩放），够用且不引入任何第三方库——
/// 这一项就把发布体积里 8.8MB 的 ClosedXML 全家桶省掉了。
/// </summary>
internal sealed class XlsxWriter
{
    private const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string NsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

    private readonly List<string> _fonts = new();
    private readonly List<string> _fills = new();
    private readonly List<string> _borders = new();
    private readonly List<string> _numFmts = new();
    private readonly List<string> _xfs = new();

    private readonly Dictionary<string, int> _fontIdx = new();
    private readonly Dictionary<string, int> _fillIdx = new();
    private readonly Dictionary<string, int> _borderIdx = new();
    private readonly Dictionary<string, int> _numFmtIdx = new();
    private readonly Dictionary<string, int> _xfIdx = new();

    private readonly string _sheetName;
    private readonly SortedDictionary<int, SortedDictionary<int, Cell>> _rows = new();
    private readonly Dictionary<int, double> _rowHeights = new();
    private readonly Dictionary<int, double> _colWidths = new();
    private readonly List<string> _merges = new();

    private int _freezeRows;

    public XlsxWriter(string sheetName)
    {
        _sheetName = sheetName;

        // fill 0 / fill 1 是 Excel 规定的占位（none 与 gray125），必须存在
        _fills.Add("<fill><patternFill patternType=\"none\"/></fill>");
        _fills.Add("<fill><patternFill patternType=\"gray125\"/></fill>");
        _fillIdx["none"] = 0;
        _fillIdx["gray125"] = 1;

        _borders.Add("<border><left/><right/><top/><bottom/><diagonal/></border>");
        _borderIdx[BorderKey(new CellStyle().Left, new CellStyle().Right,
                            new CellStyle().Top, new CellStyle().Bottom)] = 0;
    }

    // ==================== 内容 ====================

    public void SetText(int row, int col, string? text, CellStyle style)
    {
        var cell = GetCell(row, col);
        cell.Text = text;
        cell.Number = null;
        cell.Style = style;
    }

    public void SetNumber(int row, int col, double value, CellStyle style)
    {
        var cell = GetCell(row, col);
        cell.Number = value;
        cell.Text = null;
        cell.Style = style;
    }

    public void SetRowHeight(int row, double height) => _rowHeights[row] = height;

    public void SetColumnWidth(int col, double width) => _colWidths[col] = width;

    public void FreezeRows(int count) => _freezeRows = count;

    /// <summary>合并一个矩形区域，非左上角的单元格内容会被清空（与 Excel 行为一致）。</summary>
    public void Merge(int r1, int c1, int r2, int c2)
    {
        if (r1 == r2 && c1 == c2) return;

        for (var r = r1; r <= r2; r++)
        {
            for (var c = c1; c <= c2; c++)
            {
                if (r == r1 && c == c1) continue;
                if (_rows.TryGetValue(r, out var line) && line.TryGetValue(c, out var cell))
                {
                    cell.Text = null;
                    cell.Number = null;
                }
            }
        }

        _merges.Add($"{ColName(c1)}{r1}:{ColName(c2)}{r2}");
    }

    /// <summary>给矩形区域统一描边：外圈用 outerStyle，内部用内线样式。</summary>
    public void ApplyBorders(int r1, int c1, int r2, int c2,
                             string innerColor, string outerColor)
    {
        for (var r = r1; r <= r2; r++)
        {
            for (var c = c1; c <= c2; c++)
            {
                var cell = GetCell(r, c);
                var style = cell.Style ??= new CellStyle();

                var top = r == r1 ? LineStyle.Medium : LineStyle.Thin;
                var bottom = r == r2 ? LineStyle.Medium : LineStyle.Thin;
                var left = c == c1 ? LineStyle.Medium : LineStyle.Thin;
                var right = c == c2 ? LineStyle.Medium : LineStyle.Thin;

                style.Top = new BorderSide(top, top == LineStyle.Medium ? outerColor : innerColor);
                style.Bottom = new BorderSide(bottom, bottom == LineStyle.Medium ? outerColor : innerColor);
                style.Left = new BorderSide(left, left == LineStyle.Medium ? outerColor : innerColor);
                style.Right = new BorderSide(right, right == LineStyle.Medium ? outerColor : innerColor);
            }
        }
    }

    /// <summary>只覆盖某一行的上边线（用于汇总区上方那条分隔线）。</summary>
    public void SetTopBorder(int row, int c1, int c2, string color, LineStyle line)
    {
        for (var c = c1; c <= c2; c++)
        {
            var cell = GetCell(row, c);
            var style = cell.Style ??= new CellStyle();
            style.Top = new BorderSide(line, color);
        }
    }

    // ==================== 保存 ====================

    public void Save(string path)
    {
        // 先生成工作表 XML：这一步会把用到的字体/底色/边框登记进样式表，
        // 顺序不能反，否则 styles.xml 会写出一张空的 cellXfs。
        var sheetXml = Sheet();
        var stylesXml = Styles();

        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Create);

        Add(zip, "[Content_Types].xml", ContentTypes());
        Add(zip, "_rels/.rels", RootRels());
        Add(zip, "xl/workbook.xml", Workbook());
        Add(zip, "xl/_rels/workbook.xml.rels", WorkbookRels());
        Add(zip, "xl/styles.xml", stylesXml);
        Add(zip, "xl/worksheets/sheet1.xml", sheetXml);
    }

    private static void Add(ZipArchive zip, string name, string content)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var s = entry.Open();
        using var w = new StreamWriter(s, new UTF8Encoding(false));
        w.Write(content);
    }

    private static string ContentTypes() =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
        "</Types>";

    private static string RootRels() =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        $"<Relationships xmlns=\"{NsPkgRel}\">" +
        $"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
        "</Relationships>";

    private string Workbook() =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        $"<workbook xmlns=\"{Ns}\" xmlns:r=\"{NsRel}\">" +
        "<bookViews><workbookView/></bookViews>" +
        $"<sheets><sheet name=\"{Esc(_sheetName)}\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
        "</workbook>";

    private static string WorkbookRels() =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        $"<Relationships xmlns=\"{NsPkgRel}\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
        "</Relationships>";

    private string Sheet()
    {
        var sb = new StringBuilder(8192);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<worksheet xmlns=\"{Ns}\" xmlns:r=\"{NsRel}\">");

        sb.Append("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>");
        sb.Append($"<dimension ref=\"A1:{ColName(MaxCol())}{MaxRow()}\"/>");

        // 冻结窗格
        sb.Append("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">");
        if (_freezeRows > 0)
        {
            sb.Append($"<pane ySplit=\"{_freezeRows}\" topLeftCell=\"A{_freezeRows + 1}\" activePane=\"bottomLeft\" state=\"frozen\"/>");
            sb.Append($"<selection pane=\"bottomLeft\" activeCell=\"A{_freezeRows + 1}\" sqref=\"A{_freezeRows + 1}\"/>");
        }
        else
        {
            sb.Append("<selection activeCell=\"A1\" sqref=\"A1\"/>");
        }
        sb.Append("</sheetView></sheetViews>");

        sb.Append("<sheetFormatPr defaultRowHeight=\"15\"/>");

        // 列宽
        if (_colWidths.Count > 0)
        {
            sb.Append("<cols>");
            foreach (var kv in _colWidths.OrderBy(k => k.Key))
            {
                sb.Append($"<col min=\"{kv.Key}\" max=\"{kv.Key}\" width=\"{Num(kv.Value)}\" customWidth=\"1\"/>");
            }
            sb.Append("</cols>");
        }

        // 单元格
        sb.Append("<sheetData>");
        foreach (var rowPair in _rows)
        {
            var rowNo = rowPair.Key;
            _rowHeights.TryGetValue(rowNo, out var height);

            sb.Append($"<row r=\"{rowNo}\"");
            if (height > 0) sb.Append($" ht=\"{Num(height)}\" customHeight=\"1\"");
            sb.Append('>');

            foreach (var colPair in rowPair.Value)
            {
                var (colNo, cell) = (colPair.Key, colPair.Value);
                if (cell.Text == null && cell.Number == null) continue;

                var styleId = ResolveStyle(cell.Style);
                var reference = ColName(colNo) + rowNo;

                if (cell.Number is { } number)
                {
                    sb.Append($"<c r=\"{reference}\" s=\"{styleId}\"><v>{Num(number)}</v></c>");
                }
                else
                {
                    var text = cell.Text ?? string.Empty;
                    sb.Append($"<c r=\"{reference}\" s=\"{styleId}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                    sb.Append(EscLine(text));
                    sb.Append("</t></is></c>");
                }
            }

            sb.Append("</row>");
        }
        sb.Append("</sheetData>");

        // 合并
        if (_merges.Count > 0)
        {
            sb.Append($"<mergeCells count=\"{_merges.Count}\">");
            foreach (var m in _merges) sb.Append($"<mergeCell ref=\"{m}\"/>");
            sb.Append("</mergeCells>");
        }

        sb.Append("<pageMargins left=\"0.5\" right=\"0.5\" top=\"0.6\" bottom=\"0.6\" header=\"0.3\" footer=\"0.3\"/>");
        sb.Append("<pageSetup paperSize=\"9\" orientation=\"landscape\" fitToWidth=\"1\" fitToHeight=\"0\"/>");
        sb.Append("</worksheet>");
        return sb.ToString();
    }

    private int MaxRow()
    {
        var max = 1;
        foreach (var r in _rows.Keys) if (r > max) max = r;
        return max;
    }

    private int MaxCol()
    {
        var max = 1;
        foreach (var line in _rows.Values)
            foreach (var c in line.Keys)
                if (c > max) max = c;
        return max;
    }

    // ==================== 样式表 ====================

    private int ResolveStyle(CellStyle? style)
    {
        style ??= new CellStyle();

        var fontId = InternFont(style);
        var fillId = InternFill(style);
        var borderId = InternBorder(style);
        var numFmtId = InternNumberFormat(style);

        var key = $"{numFmtId}|{fontId}|{fillId}|{borderId}|{style.HAlign}|{style.VAlign}|{style.Wrap}|{style.Indent}";
        if (_xfIdx.TryGetValue(key, out var cached)) return cached;

        var sb = new StringBuilder("<xf numFmtId=\"").Append(numFmtId)
            .Append("\" fontId=\"").Append(fontId)
            .Append("\" fillId=\"").Append(fillId)
            .Append("\" borderId=\"").Append(borderId)
            .Append("\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"");

        if (numFmtId != 0) sb.Append(" applyNumberFormat=\"1\"");

        var hasAlign = style.HAlign != HAlign.General || style.VAlign != VAlign.Bottom
                       || style.Wrap || style.Indent > 0;

        if (hasAlign)
        {
            sb.Append(" applyAlignment=\"1\"><alignment");
            if (style.HAlign != HAlign.General) sb.Append(" horizontal=\"").Append(style.HAlign.ToString().ToLowerInvariant()).Append('"');
            if (style.VAlign != VAlign.Bottom) sb.Append(" vertical=\"").Append(style.VAlign.ToString().ToLowerInvariant()).Append('"');
            if (style.Wrap) sb.Append(" wrapText=\"1\"");
            if (style.Indent > 0) sb.Append(" indent=\"").Append(style.Indent).Append('"');
            sb.Append("/></xf>");
        }
        else
        {
            sb.Append("/>");
        }

        _xfs.Add(sb.ToString());
        var id = _xfs.Count;
        _xfIdx[key] = id;
        return id;
    }

    private int InternFont(CellStyle s)
    {
        var key = $"{s.FontName}|{Num(s.FontSize)}|{s.Bold}|{s.FontColor}";
        if (_fontIdx.TryGetValue(key, out var i)) return i;

        var sb = new StringBuilder("<font>");
        if (s.Bold) sb.Append("<b/>");
        sb.Append("<sz val=\"").Append(Num(s.FontSize)).Append("\"/>");
        if (!string.IsNullOrEmpty(s.FontColor)) sb.Append("<color rgb=\"").Append(s.FontColor).Append("\"/>");
        sb.Append("<name val=\"").Append(Esc(s.FontName)).Append("\"/>");
        sb.Append("<charset val=\"134\"/>");   // 134 = 简体中文
        sb.Append("</font>");

        _fonts.Add(sb.ToString());
        var id = _fonts.Count - 1;
        _fontIdx[key] = id;
        return id;
    }

    private int InternFill(CellStyle s)
    {
        if (string.IsNullOrEmpty(s.Fill)) return 0;

        if (_fillIdx.TryGetValue(s.Fill, out var i)) return i;

        _fills.Add($"<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{s.Fill}\"/><bgColor indexed=\"64\"/></patternFill></fill>");
        var id = _fills.Count - 1;
        _fillIdx[s.Fill] = id;
        return id;
    }

    private int InternBorder(CellStyle s)
    {
        var key = BorderKey(s.Left, s.Right, s.Top, s.Bottom);
        if (_borderIdx.TryGetValue(key, out var i)) return i;

        var sb = new StringBuilder("<border>");
        sb.Append(Side("left", s.Left));
        sb.Append(Side("right", s.Right));
        sb.Append(Side("top", s.Top));
        sb.Append(Side("bottom", s.Bottom));
        sb.Append("<diagonal/></border>");

        _borders.Add(sb.ToString());
        var id = _borders.Count - 1;
        _borderIdx[key] = id;
        return id;
    }

    private static string Side(string name, BorderSide side)
    {
        if (side.Style == LineStyle.None) return $"<{name}/>";
        var style = side.Style == LineStyle.Medium ? "medium" : "thin";
        return $"<{name} style=\"{style}\"><color rgb=\"{side.Color}\"/></{name}>";
    }

    private static string BorderKey(BorderSide l, BorderSide r, BorderSide t, BorderSide b) =>
        $"{(int)l.Style}{l.Color}|{(int)r.Style}{r.Color}|{(int)t.Style}{t.Color}|{(int)b.Style}{b.Color}";

    private int InternNumberFormat(CellStyle s)
    {
        if (string.IsNullOrEmpty(s.NumberFormat)) return 0;
        if (s.NumberFormat == "@") return 49;   // 文本，Excel 内置

        if (_numFmtIdx.TryGetValue(s.NumberFormat, out var i)) return i;

        var id = 164 + _numFmts.Count;
        _numFmts.Add($"<numFmt numFmtId=\"{id}\" formatCode=\"{Esc(s.NumberFormat)}\"/>");
        _numFmtIdx[s.NumberFormat] = id;
        return id;
    }

    private string Styles()
    {
        // 至少保证默认字体存在，Excel 打开时不会抱怨
        if (_fonts.Count == 0) InternFont(new CellStyle());

        var sb = new StringBuilder(4096);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<styleSheet xmlns=\"{Ns}\">");

        if (_numFmts.Count > 0)
        {
            sb.Append($"<numFmts count=\"{_numFmts.Count}\">");
            foreach (var f in _numFmts) sb.Append(f);
            sb.Append("</numFmts>");
        }

        sb.Append($"<fonts count=\"{_fonts.Count}\">");
        foreach (var f in _fonts) sb.Append(f);
        sb.Append("</fonts>");

        sb.Append($"<fills count=\"{_fills.Count}\">");
        foreach (var f in _fills) sb.Append(f);
        sb.Append("</fills>");

        sb.Append($"<borders count=\"{_borders.Count}\">");
        foreach (var b in _borders) sb.Append(b);
        sb.Append("</borders>");

        sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

        sb.Append($"<cellXfs count=\"{_xfs.Count + 1}\">");
        sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
        foreach (var x in _xfs) sb.Append(x);
        sb.Append("</cellXfs>");

        sb.Append("<cellStyles count=\"1\"><cellStyle name=\"常规\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
        sb.Append("</styleSheet>");
        return sb.ToString();
    }

    // ==================== 小工具 ====================

    private Cell GetCell(int row, int col)
    {
        if (!_rows.TryGetValue(row, out var line))
        {
            line = new SortedDictionary<int, Cell>();
            _rows[row] = line;
        }

        if (!line.TryGetValue(col, out var cell))
        {
            cell = new Cell();
            line[col] = cell;
        }

        return cell;
    }

    private static string ColName(int col)
    {
        var text = string.Empty;
        while (col > 0)
        {
            var m = (col - 1) % 26;
            text = (char)('A' + m) + text;
            col = (col - 1) / 26;
        }
        return text;
    }

    private static string Num(double value) => value.ToString("0.##########", CultureInfo.InvariantCulture);

    private static string Esc(string? text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                   .Replace("\"", "&quot;").Replace("'", "&apos;");
    }

    /// <summary>单元格文本：转义并把换行写成字符实体，配合 xml:space="preserve" 保留多行。</summary>
    private static string EscLine(string text) =>
        Esc(text.Replace("\r\n", "\n").Replace("\r", "\n")).Replace("\n", "&#10;");

    private sealed class Cell
    {
        public string? Text;
        public double? Number;
        public CellStyle? Style;
    }
}

internal enum HAlign { General, Left, Center, Right }
internal enum VAlign { Bottom, Center, Top }
internal enum LineStyle { None, Thin, Medium }

internal struct BorderSide : IEquatable<BorderSide>
{
    public LineStyle Style;
    public string Color;

    public BorderSide(LineStyle style, string color)
    {
        Style = style;
        Color = color;
    }

    public bool Equals(BorderSide other) =>
        Style == other.Style && string.Equals(Color, other.Color, StringComparison.Ordinal);

    public override bool Equals(object? obj) => obj is BorderSide other && Equals(other);

    public override int GetHashCode() => ((int)Style * 397) ^ (Color?.GetHashCode() ?? 0);
}

/// <summary>一个单元格的显示属性。未显式给的项走 Excel 默认值。</summary>
internal sealed class CellStyle
{
    public string FontName = "宋体";
    public double FontSize = 10;
    public bool Bold;
    public string? FontColor;     // ARGB，如 FF33475B
    public string? Fill;          // ARGB
    public string? NumberFormat;  // null=常规，"@"=文本，"0.00"=两位小数
    public HAlign HAlign = HAlign.General;
    public VAlign VAlign = VAlign.Bottom;
    public bool Wrap;
    public int Indent;
    public BorderSide Left;
    public BorderSide Right;
    public BorderSide Top;
    public BorderSide Bottom;
}
