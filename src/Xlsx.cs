using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

class SheetInfo
{
    public int Offset;
    public string Name;
    public string RId;
    public string Path;

    public SheetInfo(string name, string rid)
    {
        Name = name;
        RId = rid;
    }
}

delegate void CellSink(string sheetName, int row, int col, string value, string type, string formula, int styleIndex);

class DefinedName
{
    public string Name;
    public int LocalSheetId = -1;
    public bool Hidden;
    public string Comment;
    public string Ref;
}

class Xlsx : IBook
{
    const string REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    ZipArchive zip;
    Dictionary<string, ZipArchiveEntry> entries;
    List<SheetInfo> sheets = new List<SheetInfo>();
    string[] sharedStrings = new string[0];
    string[] cellFormats = new string[0];
    public bool Date1904 { get; set; }
    List<DefinedName> definedNames = new List<DefinedName>();
    List<string> allSheetNames = new List<string>();
    XmlReaderSettings settings;

    public Xlsx(string path)
    {
        settings = new XmlReaderSettings();
        settings.IgnoreWhitespace = true;
        settings.IgnoreComments = true;
        settings.DtdProcessing = DtdProcessing.Prohibit;
        settings.CloseInput = true;
        settings.CheckCharacters = false;

        zip = ZipFile.OpenRead(path);
        entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (ZipArchiveEntry e in zip.Entries)
        {
            if (e.FullName.EndsWith("/")) continue;
            entries[e.FullName] = e;
        }

        LoadWorkbook();
        LoadSharedStrings();
        LoadStyles();
    }

    public List<SheetInfo> Sheets { get { return sheets; } }

    public bool HasVba
    {
        get { return entries.ContainsKey("xl/vbaProject.bin"); }
    }

    public byte[] GetVbaProject()
    {
        ZipArchiveEntry e;
        if (!entries.TryGetValue("xl/vbaProject.bin", out e)) return null;
        using (Stream s = e.Open())
        {
            MemoryStream ms = new MemoryStream();
            byte[] buf = new byte[8192];
            int n;
            while ((n = s.Read(buf, 0, buf.Length)) > 0) ms.Write(buf, 0, n);
            return ms.ToArray();
        }
    }

    Stream OpenEntry(string name)
    {
        ZipArchiveEntry e;
        if (!entries.TryGetValue(name, out e)) throw new Exception("missing part: " + name);
        return e.Open();
    }

    // ------------------------------------------------------------ workbook

    void LoadWorkbook()
    {
        Dictionary<string, string> rels = new Dictionary<string, string>(StringComparer.Ordinal);

        using (Stream s = OpenEntry("xl/workbook.xml"))
        using (XmlReader r = XmlReader.Create(s, settings))
        {
            r.Read();
            while (!r.EOF)
            {
                if (r.NodeType == XmlNodeType.Element)
                {
                    string ln = r.LocalName;

                    if (ln == "workbookPr")
                    {
                        string d = r.GetAttribute("date1904");
                        if (d == "1" || string.Equals(d, "true", StringComparison.OrdinalIgnoreCase)) Date1904 = true;
                        r.Read();
                        continue;
                    }

                    if (ln == "sheet")
                    {
                        string name = r.GetAttribute("name");
                        string rid = r.GetAttribute("id", REL_NS);
                        sheets.Add(new SheetInfo(name, rid));
                        r.Read();
                        continue;
                    }

                    if (ln == "definedName")
                    {
                        DefinedName dn = new DefinedName();
                        dn.Name = r.GetAttribute("name");
                        string lsid = r.GetAttribute("localSheetId");
                        int sid;
                        if (lsid != null && int.TryParse(lsid, NumberStyles.Integer, CultureInfo.InvariantCulture, out sid))
                            dn.LocalSheetId = sid;
                        string hid = r.GetAttribute("hidden");
                        dn.Hidden = (hid == "1" || string.Equals(hid, "true", StringComparison.OrdinalIgnoreCase));
                        dn.Comment = r.GetAttribute("comment");
                        dn.Ref = r.ReadElementContentAsString();
                        definedNames.Add(dn);
                        continue; // ReadElementContentAsString already advanced
                    }
                }
                r.Read();
            }
        }

        if (entries.ContainsKey("xl/_rels/workbook.xml.rels"))
        {
            using (Stream s = OpenEntry("xl/_rels/workbook.xml.rels"))
            using (XmlReader r = XmlReader.Create(s, settings))
            {
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship")
                    {
                        string id = r.GetAttribute("Id");
                        string target = r.GetAttribute("Target");
                        string mode = r.GetAttribute("TargetMode");
                        if (id != null && target != null && mode != "External")
                            rels[id] = target;
                    }
                }
            }
        }

        for (int i = 0; i < sheets.Count; i++)
        {
            SheetInfo sh = sheets[i];
            string target;
            if (sh.RId != null && rels.TryGetValue(sh.RId, out target))
                sh.Path = Resolve("xl/", target);
        }

        allSheetNames = new List<string>();
        for (int i = 0; i < sheets.Count; i++) allSheetNames.Add(sheets[i].Name);

        // Drop sheets we cannot read (chartsheets, missing parts).
        List<SheetInfo> keep = new List<SheetInfo>();
        for (int i = 0; i < sheets.Count; i++)
        {
            if (sheets[i].Path != null && entries.ContainsKey(sheets[i].Path)) keep.Add(sheets[i]);
        }
        sheets = keep;
    }

    static string Resolve(string baseDir, string target)
    {
        if (target.StartsWith("/")) return target.TrimStart('/');

        string combined = baseDir + target;
        string[] parts = combined.Split('/');
        List<string> stack = new List<string>();
        for (int i = 0; i < parts.Length; i++)
        {
            string p = parts[i];
            if (p.Length == 0 || p == ".") continue;
            if (p == "..")
            {
                if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
            }
            else
            {
                stack.Add(p);
            }
        }
        return string.Join("/", stack.ToArray());
    }

    // ------------------------------------------------------------ strings

    void LoadSharedStrings()
    {
        if (!entries.ContainsKey("xl/sharedStrings.xml")) return;

        List<string> list = new List<string>();
        using (Stream s = OpenEntry("xl/sharedStrings.xml"))
        using (XmlReader r = XmlReader.Create(s, settings))
        {
            while (r.Read())
            {
                if (r.NodeType == XmlNodeType.Element && r.LocalName == "si")
                {
                    if (r.IsEmptyElement) { list.Add(""); continue; }
                    list.Add(ReadSi(r));
                }
            }
        }
        sharedStrings = list.ToArray();
    }

    static string ReadSi(XmlReader r)
    {
        StringBuilder sb = new StringBuilder();
        int depth = r.Depth;
        while (!r.EOF)
        {
            if (r.NodeType == XmlNodeType.EndElement && r.Depth == depth) return sb.ToString();
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "t")
            {
                sb.Append(r.ReadElementContentAsString());
                continue;
            }
            r.Read();
        }
        return sb.ToString();
    }

    // ------------------------------------------------------------ cells

    public string GetDimension(SheetInfo sh)
    {
        using (Stream s = OpenEntry(sh.Path))
        using (XmlReader r = XmlReader.Create(s, settings))
        {
            while (r.Read())
            {
                if (r.NodeType != XmlNodeType.Element) continue;
                if (r.LocalName == "dimension") return r.GetAttribute("ref");
                if (r.LocalName == "sheetData") break;
            }
        }
        return null;
    }

    public void ReadCells(SheetInfo sheet, int r1, int c1, int r2, int c2, CellSink sink)
    {
        using (Stream s = OpenEntry(sheet.Path))
        using (XmlReader r = XmlReader.Create(s, settings))
        {
            int rowNum = 0;
            int lastCol = 0;

            r.Read();
            while (!r.EOF)
            {
                if (r.NodeType == XmlNodeType.Element)
                {
                    string ln = r.LocalName;

                    if (ln == "row")
                    {
                        string rs = r.GetAttribute("r");
                        if (rs != null) rowNum = int.Parse(rs, CultureInfo.InvariantCulture);
                        else rowNum++;
                        lastCol = 0;
                        if (r2 > 0 && rowNum > r2) break;
                        r.Read();
                        continue;
                    }

                    if (ln == "c")
                    {
                        string refAttr = r.GetAttribute("r");
                        string tAttr = r.GetAttribute("t");
                        string sAttr = r.GetAttribute("s");
                        int styleIndex = 0;
                        if (sAttr != null) int.TryParse(sAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out styleIndex);

                        int row, col;
                        if (refAttr != null)
                        {
                            if (!A1.TryParseRef(refAttr, out row, out col)) { r.Skip(); continue; }
                        }
                        else
                        {
                            col = lastCol + 1;
                            row = rowNum;
                        }
                        lastCol = col;

                        bool inRange = (row >= r1 && row <= r2 && col >= c1 && col <= c2);

                        if (r.IsEmptyElement)
                        {
                            r.Read();
                            if (inRange) sink(sheet.Name, row, col, null, null, null, styleIndex);
                            continue;
                        }

                        if (!inRange)
                        {
                            r.Skip();
                            continue;
                        }

                        string v = null;
                        string isText = null;
                        string f = null;
                        int depth = r.Depth;

                        while (!r.EOF)
                        {
                            if (r.NodeType == XmlNodeType.EndElement && r.Depth == depth) break;
                            if (r.NodeType == XmlNodeType.Element)
                            {
                                string ln2 = r.LocalName;
                                if (ln2 == "v") { v = r.ReadElementContentAsString(); continue; }
                                if (ln2 == "f") { f = r.ReadElementContentAsString(); continue; }
                                if (ln2 == "is")
                                {
                                    isText = ReadSi(r);
                                    if (r.IsEmptyElement) r.Read();
                                    continue;
                                }
                            }
                            r.Read();
                        }

                        string val;
                        string type;
                        CellValue(tAttr, v, isText, out val, out type);
                        sink(sheet.Name, row, col, val, type, f, styleIndex);
                        continue;
                    }

                    if (ln == "dimension") { r.Skip(); continue; }
                }
                r.Read();
            }
        }
    }

    void CellValue(string tAttr, string v, string isText, out string val, out string type)
    {
        val = null;
        type = null;

        if (tAttr == "s")
        {
            int idx;
            if (v != null && int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx)
                && idx >= 0 && idx < sharedStrings.Length)
            {
                val = sharedStrings[idx];
                type = "s";
            }
            else if (v != null)
            {
                val = v;
                type = "s";
            }
            return;
        }

        if (tAttr == "inlineStr")
        {
            if (isText != null) { val = isText; type = "s"; }
            return;
        }

        if (tAttr == "str") { if (v != null) { val = v; type = "s"; } return; }
        if (tAttr == "b") { if (v != null) { val = (v == "1") ? "true" : "false"; type = "b"; } return; }
        if (tAttr == "e") { if (v != null) { val = v; type = "e"; } return; }
        if (tAttr == "d") { if (v != null) { val = v; type = "d"; } return; }

        if (v != null && v.Length > 0)
        {
            if (IsJsonNumber(v)) { val = v; type = "n"; }
            else { val = v; type = "s"; }
        }
        else if (isText != null)
        {
            val = isText;
            type = "s";
        }
    }

    static bool IsJsonNumber(string s)
    {
        int i = 0;
        if (s[i] == '-') i++;
        int digits = 0;
        while (i < s.Length && s[i] >= '0' && s[i] <= '9') { i++; digits++; }
        if (i < s.Length && s[i] == '.')
        {
            i++;
            while (i < s.Length && s[i] >= '0' && s[i] <= '9') { i++; digits++; }
        }
        if (digits == 0) return false;
        if (i < s.Length && (s[i] == 'e' || s[i] == 'E'))
        {
            i++;
            if (i < s.Length && (s[i] == '+' || s[i] == '-')) i++;
            int ed = 0;
            while (i < s.Length && s[i] >= '0' && s[i] <= '9') { i++; ed++; }
            if (ed == 0) return false;
        }
        return i == s.Length;
    }

    // ------------------------------------------------------------ styles

    void LoadStyles()
    {
        Dictionary<int, string> custom = new Dictionary<int, string>();
        List<int> ids = new List<int>();

        if (entries.ContainsKey("xl/styles.xml"))
        {
            using (Stream s = OpenEntry("xl/styles.xml"))
            using (XmlReader r = XmlReader.Create(s, settings))
            {
                bool inCellXfs = false;
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.Element)
                    {
                        string ln = r.LocalName;
                        if (ln == "numFmt")
                        {
                            string idS = r.GetAttribute("numFmtId");
                            string code = r.GetAttribute("formatCode");
                            int id;
                            if (idS != null && code != null && int.TryParse(idS, NumberStyles.Integer, CultureInfo.InvariantCulture, out id))
                                custom[id] = code;
                        }
                        else if (ln == "cellXfs") inCellXfs = true;
                        else if (ln == "xf" && inCellXfs)
                        {
                            string idS = r.GetAttribute("numFmtId");
                            int id = 0;
                            if (idS != null) int.TryParse(idS, NumberStyles.Integer, CultureInfo.InvariantCulture, out id);
                            ids.Add(id);
                        }
                    }
                    else if (r.NodeType == XmlNodeType.EndElement && r.LocalName == "cellXfs")
                    {
                        inCellXfs = false;
                    }
                }
            }
        }

        cellFormats = new string[ids.Count];
        for (int i = 0; i < ids.Count; i++)
        {
            string code;
            if (custom.TryGetValue(ids[i], out code)) cellFormats[i] = code;
            else cellFormats[i] = BuiltinFormat(ids[i]);
        }
    }

    public string GetFormatCode(int styleIndex)
    {
        if (styleIndex < 0 || styleIndex >= cellFormats.Length) return "General";
        return cellFormats[styleIndex];
    }

    internal static string BuiltinFormat(int id)
    {
        switch (id)
        {
            case 0: return "General";
            case 1: return "0";
            case 2: return "0.00";
            case 3: return "#,##0";
            case 4: return "#,##0.00";
            case 5: return "\"$\"#,##0_);(\"$\"#,##0)";
            case 6: return "\"$\"#,##0_);[Red](\"$\"#,##0)";
            case 7: return "\"$\"#,##0.00_);(\"$\"#,##0.00)";
            case 8: return "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
            case 9: return "0%";
            case 10: return "0.00%";
            case 11: return "0.00E+00";
            case 12: return "# ?/?";
            case 13: return "# ??/??";
            case 14: return "mm-dd-yy";
            case 15: return "d-mmm-yy";
            case 16: return "d-mmm";
            case 17: return "mmm-yy";
            case 18: return "h:mm AM/PM";
            case 19: return "h:mm:ss AM/PM";
            case 20: return "h:mm";
            case 21: return "h:mm:ss";
            case 22: return "m/d/yy h:mm";
            case 37: return "#,##0 ;(#,##0)";
            case 38: return "#,##0 ;[Red](#,##0)";
            case 39: return "#,##0.00;(#,##0.00)";
            case 40: return "#,##0.00;[Red](#,##0.00)";
            case 45: return "mm:ss";
            case 46: return "[h]:mm:ss";
            case 47: return "mmss.0";
            case 48: return "##0.0E+0";
            case 49: return "@";
            default: return "General";
        }
    }

    public List<DefinedName> DefinedNames { get { return definedNames; } }

    public string SheetName(int id)
    {
        if (id >= 0 && id < allSheetNames.Count) return allSheetNames[id];
        return null;
    }

    public DefinedName FindName(string name, string sheetName)
    {
        DefinedName workbookScope = null;
        DefinedName localScope = null;
        DefinedName anyLocal = null;

        for (int i = 0; i < definedNames.Count; i++)
        {
            DefinedName dn = definedNames[i];
            if (!string.Equals(dn.Name, name, StringComparison.OrdinalIgnoreCase)) continue;

            if (dn.LocalSheetId < 0)
            {
                if (workbookScope == null) workbookScope = dn;
            }
            else
            {
                if (anyLocal == null) anyLocal = dn;
                if (sheetName != null && string.Equals(SheetName(dn.LocalSheetId), sheetName, StringComparison.Ordinal))
                    localScope = dn;
            }
        }

        if (localScope != null) return localScope;
        if (workbookScope != null) return workbookScope;
        return anyLocal;
    }

    public bool TryResolveNameRange(string name, string sheetName,
        out string targetSheet, out int r1, out int c1, out int r2, out int c2)
    {
        targetSheet = null;
        r1 = 1; c1 = 1; r2 = 0; c2 = 0;

        DefinedName dn = FindName(name, sheetName);
        if (dn == null) return false;
        if (!ParseNameRef(dn, sheetName, out targetSheet, out r1, out c1, out r2, out c2)) return false;
        if (targetSheet == null && dn.LocalSheetId >= 0) targetSheet = SheetName(dn.LocalSheetId);
        return true;
    }

    static bool ParseNameRef(DefinedName dn, string fallbackSheet,
        out string targetSheet, out int r1, out int c1, out int r2, out int c2)
    {
        targetSheet = null;
        r1 = 1; c1 = 1; r2 = 0; c2 = 0;

        if (dn.Ref == null) return false;
        string text = dn.Ref.Trim();
        if (text.StartsWith("=")) text = text.Substring(1).Trim();
        if (text.Length == 0) return false;

        // unions / multi-area names cannot be expressed as one rectangle
        if (HasTopLevelComma(text)) return false;

        string rangePart = text;
        int bang = FindBang(text);
        if (bang >= 0)
        {
            targetSheet = UnquoteSheet(text.Substring(0, bang));
            rangePart = text.Substring(bang + 1).Trim();
        }
        else if (dn.LocalSheetId >= 0)
        {
            targetSheet = null; // filled below by caller via SheetName
        }
        else
        {
            targetSheet = fallbackSheet;
        }

        if (!A1.TryParseRange(rangePart, out r1, out c1, out r2, out c2)) return false;
        return true;
    }

    static bool HasTopLevelComma(string s)
    {
        bool inQuote = false;
        int depth = 0;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == '\'') { inQuote = !inQuote; continue; }
            if (inQuote) continue;
            if (c == '(') depth++;
            else if (c == ')') { if (depth > 0) depth--; }
            else if (c == ',' && depth == 0) return true;
        }
        return false;
    }

    static int FindBang(string s)
    {
        bool inQuote = false;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == '\'') { inQuote = !inQuote; continue; }
            if (!inQuote && c == '!') return i;
        }
        return -1;
    }

    static string UnquoteSheet(string s)
    {
        s = s.Trim();
        if (s.Length >= 2 && s[0] == '\'' && s[s.Length - 1] == '\'')
            return s.Substring(1, s.Length - 2).Replace("''", "'");
        return s;
    }

    public void Dispose()
    {
        if (zip != null) { zip.Dispose(); zip = null; }
    }
}
