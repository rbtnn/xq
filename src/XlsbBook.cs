using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

class SharedRange
{
    public int R1, C1, R2, C2;
    public byte[] Tokens;
}

class XlsbSheet
{
    public SheetInfo Info;
    public int R1 = -1, C1 = -1, R2 = -1, C2 = -1;
    public List<SharedRange> Shared = new List<SharedRange>();
}

class NameRaw12
{
    public DefinedName Dn;
    public byte[] Tokens;
}

// BIFF12 (.xlsb) reader. ZIP package + binary parts.
class XlsbBook : BookBase
{
    ZipArchive zip;
    List<SheetInfo> sheets = new List<SheetInfo>();
    List<XlsbSheet> xlsbSheets = new List<XlsbSheet>();
    string[] sst = new string[0];
    List<int> xfFmt = new List<int>();
    Dictionary<int, string> formats = new Dictionary<int, string>();
    bool date1904;
    List<string> externSheets = new List<string>();
    List<NameRaw12> rawNames = new List<NameRaw12>();
    Dictionary<string, string> rels = new Dictionary<string, string>();
    Dictionary<string, string> relTypes = new Dictionary<string, string>();

    public XlsbBook(string path)
    {
        zip = ZipFile.OpenRead(path);
        ReadRels();
        ReadWorkbook();
        ResolveNames();
        ReadStyles();
        ReadSharedStrings();
        ReadSheets();
    }

    public override List<SheetInfo> Sheets { get { return sheets; } }
    public override bool Date1904 { get { return date1904; } }
    public override bool HasVba { get { return zip.GetEntry("xl/vbaProject.bin") != null; } }

    public override byte[] GetVbaProject()
    {
        return ReadEntry("xl/vbaProject.bin");
    }

    public override void Dispose()
    {
        zip.Dispose();
    }

    // ------------------------------------------------------------ package

    byte[] ReadEntry(string name)
    {
        ZipArchiveEntry e = zip.GetEntry(name);
        if (e == null) return null;
        using (Stream s = e.Open())
        {
            MemoryStream ms = new MemoryStream();
            s.CopyTo(ms);
            return ms.ToArray();
        }
    }

    void ReadRels()
    {
        byte[] data = ReadEntry("xl/_rels/workbook.bin.rels");
        if (data == null) return;

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(Encoding.UTF8.GetString(data));
        XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

        XmlNodeList nodes = doc.SelectNodes("//r:Relationship", ns);
        for (int i = 0; i < nodes.Count; i++)
        {
            XmlElement e = nodes[i] as XmlElement;
            if (e == null) continue;
            string id = e.GetAttribute("Id");
            string target = e.GetAttribute("Target");
            string type = e.GetAttribute("Type");
            string mode = e.GetAttribute("TargetMode");

            if (mode == "External" || id.Length == 0 || target.Length == 0) continue;

            if (target.StartsWith("/")) target = target.Substring(1);
            else target = "xl/" + target;

            rels[id] = target;
            relTypes[id] = type;
        }
    }

    // ------------------------------------------------------------ workbook

    void ReadWorkbook()
    {
        byte[] wb = ReadEntry("xl/workbook.bin");
        if (wb == null) throw new Exception("xl/workbook.bin not found");

        Biff12Reader r = new Biff12Reader(wb);
        int id, off, len;
        while (r.Next(out id, out off, out len))
        {
            if (id == 0x0099) // BrtWbProp
            {
                date1904 = (wb[off] & 1) != 0;
                continue;
            }

            if (id == 0x009C) // BrtBundleSh
            {
                int visible = (int)ReadU32(wb, off);
                int sheetId = (int)ReadU32(wb, off + 4);
                uint relLen = ReadU32(wb, off + 8);
                string relId = null;
                int namePos;
                if (relLen != 0xFFFFFFFF)
                {
                    relId = Encoding.Unicode.GetString(wb, off + 12, (int)relLen * 2);
                    namePos = off + 12 + (int)relLen * 2;
                }
                else
                {
                    namePos = off + 12;
                }

                int next;
                string name = ReadWideString(wb, namePos, out next);

                SheetInfo sh = new SheetInfo(name, relId);
                if (relId != null)
                {
                    string path;
                    if (rels.TryGetValue(relId, out path)) sh.Path = path;
                }
                sheets.Add(sh);
                continue;
            }

            if (id == 0x016A) // BrtExternSheet
            {
                int cxti = (int)ReadU32(wb, off);
                for (int i = 0; i < cxti; i++)
                {
                    int o = off + 4 + i * 12;
                    int sup = (int)ReadU32(wb, o);
                    int first = (int)ReadU32(wb, o + 4);
                    int last = (int)ReadU32(wb, o + 8);
                    string s = ItabName(first);
                    externSheets.Add(s);
                }
                continue;
            }

            if (id == 0x0027) // BrtName
            {
                ParseName(wb, off, len);
                continue;
            }
        }

        for (int i = 0; i < sheets.Count; i++) allSheetNames.Add(sheets[i].Name);
    }

    string ItabName(int itab)
    {
        if (itab == -2) return "#ThisWorkbook";
        if (itab == -1) return "#InvalidWorkSheet";
        if (itab >= 0 && itab < sheets.Count) return sheets[itab].Name;
        return "#Unknown";
    }

    void ParseName(byte[] wb, int off, int len)
    {
        int flags = (int)ReadU32(wb, off);
        int itab = (int)ReadU32(wb, off + 5);

        int next;
        string name = ReadWideString(wb, off + 9, out next);

        int cce = (int)ReadU32(wb, next);
        byte[] tokens = Slice(wb, next + 4, cce);

        DefinedName dn = new DefinedName();
        dn.Name = name;
        dn.LocalSheetId = (itab == -1) ? -1 : itab;
        dn.Hidden = (flags & 1) != 0;
        dn.Ref = "";
        definedNames.Add(dn);

        NameRaw12 raw = new NameRaw12();
        raw.Dn = dn;
        raw.Tokens = tokens;
        rawNames.Add(raw);
    }

    void ResolveNames()
    {
        string[] names = allSheetNames.ToArray();
        for (int i = 0; i < rawNames.Count; i++)
        {
            NameRaw12 raw = rawNames[i];
            FormulaContext ctx = new FormulaContext();
            ctx.SheetNames = names;
            ctx.Names = definedNames;
            ctx.ExternSheet = ExternSheet;
            ctx.RelDelta = true;
            ctx.BaseRow = 0;
            ctx.BaseCol = 0;
            raw.Dn.Ref = Ptg12.Decompile(raw.Tokens, raw.Tokens.Length, ctx);
        }
    }

    string ExternSheet(int ixti)
    {
        if (ixti < 0 || ixti >= externSheets.Count) return null;
        return externSheets[ixti];
    }

    // ------------------------------------------------------------ styles

    void ReadStyles()
    {
        byte[] data = ReadEntry("xl/styles.bin");
        if (data == null) return;

        Biff12Reader r = new Biff12Reader(data);
        int id, off, len;
        while (r.Next(out id, out off, out len))
        {
            if (id == 0x0267) // BrtBeginFmts
            {
                int count = (int)ReadU32(data, off);
                for (int i = 0; i < count; i++)
                {
                    int id2, off2, len2;
                    if (!r.Next(out id2, out off2, out len2)) break;
                    if (id2 != 0x002C) continue; // BrtFmt
                    int ifmt = ReadU16(data, off2);
                    int next;
                    string code = ReadWideString(data, off2 + 2, out next);
                    formats[ifmt] = code;
                }
                continue;
            }

            if (id == 0x0269) // BrtBeginCellXFs
            {
                int count = (int)ReadU32(data, off);
                for (int i = 0; i < count; i++)
                {
                    int id2, off2, len2;
                    if (!r.Next(out id2, out off2, out len2)) break;
                    if (id2 != 0x002F) continue; // BrtXF
                    xfFmt.Add(ReadU16(data, off2 + 2));
                }
                continue;
            }
        }
    }

    // ------------------------------------------------------------ strings

    void ReadSharedStrings()
    {
        byte[] data = ReadEntry("xl/sharedStrings.bin");
        if (data == null) return;

        Biff12Reader r = new Biff12Reader(data);
        int id, off, len;
        while (r.Next(out id, out off, out len))
        {
            if (id == 0x009F) // BrtBeginSst
            {
                int count = (int)ReadU32(data, off + 4);
                List<string> list = new List<string>();
                for (int i = 0; i < count; i++)
                {
                    int id2, off2, len2;
                    if (!r.Next(out id2, out off2, out len2)) break;
                    if (id2 != 0x0013) continue; // BrtSSTItem
                    int next;
                    list.Add(ReadWideString(data, off2 + 1, out next));
                }
                sst = list.ToArray();
                return;
            }
        }
    }

    // ------------------------------------------------------------ sheets

    void ReadSheets()
    {
        for (int i = 0; i < sheets.Count; i++)
        {
            SheetInfo sh = sheets[i];
            XlsbSheet xs = new XlsbSheet();
            xs.Info = sh;

            if (sh.Path != null)
            {
                byte[] data = ReadEntry(sh.Path);
                if (data != null)
                {
                    Biff12Reader r = new Biff12Reader(data);
                    int id, off, len;
                    while (r.Next(out id, out off, out len))
                    {
                        if (id == 0x0094) // BrtWsDim
                        {
                            xs.R1 = (int)ReadU32(data, off);
                            xs.R2 = (int)ReadU32(data, off + 4);
                            xs.C1 = (int)ReadU32(data, off + 8);
                            xs.C2 = (int)ReadU32(data, off + 12);
                            continue;
                        }

                        if (id == 0x01AA) // BrtArrFmla (some producers omit fAlwaysCalc)
                        {
                            SharedRange sr = new SharedRange();
                            sr.R1 = (int)ReadU32(data, off);
                            sr.R2 = (int)ReadU32(data, off + 4);
                            sr.C1 = (int)ReadU32(data, off + 8);
                            sr.C2 = (int)ReadU32(data, off + 12);
                            int cce = (int)ReadU32(data, off + 17);
                            int tok = off + 21;
                            if (cce < 0 || tok + cce > off + len)
                            {
                                cce = (int)ReadU32(data, off + 16);
                                tok = off + 20;
                            }
                            sr.Tokens = Slice(data, tok, cce);
                            xs.Shared.Add(sr);
                            continue;
                        }

                        if (id == 0x01AB) // BrtShrFmla
                        {
                            SharedRange sr = new SharedRange();
                            sr.R1 = (int)ReadU32(data, off);
                            sr.R2 = (int)ReadU32(data, off + 4);
                            sr.C1 = (int)ReadU32(data, off + 8);
                            sr.C2 = (int)ReadU32(data, off + 12);
                            int cce = (int)ReadU32(data, off + 16);
                            int tok = off + 20;
                            if (cce < 0 || tok + cce > off + len)
                            {
                                cce = (int)ReadU32(data, off + 17);
                                tok = off + 21;
                            }
                            sr.Tokens = Slice(data, tok, cce);
                            xs.Shared.Add(sr);
                            continue;
                        }
                    }
                }
            }

            xlsbSheets.Add(xs);
        }
    }

    public override string GetDimension(SheetInfo sh)
    {
        XlsbSheet xs = FindSheet(sh);
        if (xs == null || xs.R1 < 0 || xs.R2 < xs.R1) return null;
        return A1.IndexToCol(xs.C1 + 1) + (xs.R1 + 1).ToString(CultureInfo.InvariantCulture)
            + ":" + A1.IndexToCol(xs.C2 + 1) + (xs.R2 + 1).ToString(CultureInfo.InvariantCulture);
    }

    XlsbSheet FindSheet(SheetInfo sh)
    {
        for (int i = 0; i < xlsbSheets.Count; i++)
            if (xlsbSheets[i].Info == sh) return xlsbSheets[i];
        return null;
    }

    public override string GetFormatCode(int styleIndex)
    {
        if (styleIndex < 0 || styleIndex >= xfFmt.Count) return "General";
        int id = xfFmt[styleIndex];
        string code;
        if (formats.TryGetValue(id, out code)) return code;
        return Xlsx.BuiltinFormat(id);
    }

    // ------------------------------------------------------------ cells

    public override void ReadCells(SheetInfo sh, int r1, int c1, int r2, int c2, CellSink sink)
    {
        XlsbSheet xs = FindSheet(sh);
        if (xs == null || xs.Info.Path == null) return;

        byte[] data = ReadEntry(xs.Info.Path);
        if (data == null) return;

        FormulaContext ctx = new FormulaContext();
        ctx.SheetNames = allSheetNames.ToArray();
        ctx.Names = definedNames;
        ctx.ExternSheet = ExternSheet;
        ctx.RelDelta = false;
        ctx.BaseRow = 0;
        ctx.BaseCol = 0;
        ctx.SharedFormulaEx = delegate (int masterRow, int cellCol, out int baseRow, out int baseCol)
        {
            for (int i = 0; i < xs.Shared.Count; i++)
            {
                SharedRange sr = xs.Shared[i];
                if (ctx.CellRow >= sr.R1 && ctx.CellRow <= sr.R2 && cellCol >= sr.C1 && cellCol <= sr.C2)
                {
                    baseRow = sr.R1;
                    baseCol = sr.C1;
                    return sr.Tokens;
                }
            }
            baseRow = 0;
            baseCol = 0;
            return null;
        };

        Biff12Reader r = new Biff12Reader(data);
        int id, off, len;
        int row = 0;

        while (r.Next(out id, out off, out len))
        {
            if (id == 0x0000) // BrtRowHdr
            {
                row = (int)ReadU32(data, off);
                continue;
            }

            if (id == 0x0092) break; // BrtEndSheetData

            if (id >= 0x0001 && id <= 0x000B)
            {
                int col = (int)ReadU32(data, off);
                int style = (int)(ReadU32(data, off + 4) & 0xFFFFFF);

                if (!InRange(row + 1, col + 1, r1, c1, r2, c2)) continue;

                if (id == 0x0001) // BrtCellBlank
                {
                    sink(sh.Name, row + 1, col + 1, null, null, null, style);
                }
                else if (id == 0x0002) // BrtCellRk
                {
                    int rk = (int)ReadU32(data, off + 8);
                    sink(sh.Name, row + 1, col + 1, Num(Rk(rk)), "n", null, style);
                }
                else if (id == 0x0003) // BrtCellError
                {
                    sink(sh.Name, row + 1, col + 1, ErrText(data[off + 8]), "e", null, style);
                }
                else if (id == 0x0004) // BrtCellBool
                {
                    sink(sh.Name, row + 1, col + 1, data[off + 8] != 0 ? "true" : "false", "b", null, style);
                }
                else if (id == 0x0005) // BrtCellReal
                {
                    double v = BitConverter.ToDouble(data, off + 8);
                    sink(sh.Name, row + 1, col + 1, Num(v), "n", null, style);
                }
                else if (id == 0x0006) // BrtCellSt
                {
                    int next;
                    string v = ReadWideString(data, off + 8, out next);
                    sink(sh.Name, row + 1, col + 1, v, "s", null, style);
                }
                else if (id == 0x0007) // BrtCellIsst
                {
                    int isst = (int)ReadU32(data, off + 8);
                    string v = (isst >= 0 && isst < sst.Length) ? sst[isst] : "";
                    sink(sh.Name, row + 1, col + 1, v, "s", null, style);
                }
                else if (id == 0x0008) // BrtFmlaString
                {
                    int cch = (int)ReadU32(data, off + 8);
                    string v = Encoding.Unicode.GetString(data, off + 12, cch * 2);
                    int fpos = off + 14 + cch * 2;
                    int cce = (int)ReadU32(data, fpos);
                    byte[] tokens = Slice(data, fpos + 4, cce);

                    ctx.CellRow = row;
                    ctx.CellCol = col;
                    string formula = Ptg12.Decompile(tokens, tokens.Length, ctx);
                    sink(sh.Name, row + 1, col + 1, v, "s", formula, style);
                }
                else if (id == 0x0009) // BrtFmlaNum
                {
                    double v = BitConverter.ToDouble(data, off + 8);
                    int fpos = off + 18;
                    int cce = (int)ReadU32(data, fpos);
                    byte[] tokens = Slice(data, fpos + 4, cce);

                    ctx.CellRow = row;
                    ctx.CellCol = col;
                    string formula = Ptg12.Decompile(tokens, tokens.Length, ctx);
                    sink(sh.Name, row + 1, col + 1, Num(v), "n", formula, style);
                }
                else if (id == 0x000A) // BrtFmlaBool
                {
                    bool v = data[off + 8] != 0;
                    int fpos = off + 11;
                    int cce = (int)ReadU32(data, fpos);
                    byte[] tokens = Slice(data, fpos + 4, cce);

                    ctx.CellRow = row;
                    ctx.CellCol = col;
                    string formula = Ptg12.Decompile(tokens, tokens.Length, ctx);
                    sink(sh.Name, row + 1, col + 1, v ? "true" : "false", "b", formula, style);
                }
                else if (id == 0x000B) // BrtFmlaError
                {
                    int fpos = off + 11;
                    int cce = (int)ReadU32(data, fpos);
                    byte[] tokens = Slice(data, fpos + 4, cce);

                    ctx.CellRow = row;
                    ctx.CellCol = col;
                    string formula = Ptg12.Decompile(tokens, tokens.Length, ctx);
                    sink(sh.Name, row + 1, col + 1, ErrText(data[off + 8]), "e", formula, style);
                }
            }
        }
    }

    static bool InRange(int row, int col, int r1, int c1, int r2, int c2)
    {
        return row >= r1 && row <= r2 && col >= c1 && col <= c2;
    }

    static string Num(double d)
    {
        string s = d.ToString("R", CultureInfo.InvariantCulture);
        if (s.EndsWith(".0")) s = s.Substring(0, s.Length - 2);
        return s;
    }

    static double Rk(int rk)
    {
        bool fx100 = (rk & 1) != 0;
        bool fInt = (rk & 2) != 0;
        double v;
        if (fInt)
        {
            v = (double)(rk >> 2);
        }
        else
        {
            long bits = ((long)(rk & 0xFFFFFFFC)) << 32;
            v = BitConverter.Int64BitsToDouble(bits);
        }
        if (fx100) v /= 100.0;
        return v;
    }

    static string ErrText(int code)
    {
        switch (code)
        {
            case 0x00: return "#NULL!";
            case 0x07: return "#DIV/0!";
            case 0x0F: return "#VALUE!";
            case 0x17: return "#REF!";
            case 0x1D: return "#NAME?";
            case 0x24: return "#NUM!";
            case 0x2A: return "#N/A";
            case 0x2B: return "#GETTING_DATA";
            default: return "#ERR" + code.ToString(CultureInfo.InvariantCulture);
        }
    }

    // ------------------------------------------------------------ helpers

    static string ReadWideString(byte[] d, int pos, out int next)
    {
        int cch = (int)ReadU32(d, pos);
        next = pos + 4 + cch * 2;
        return Encoding.Unicode.GetString(d, pos + 4, cch * 2);
    }

    static int ReadU16(byte[] d, int p)
    {
        return d[p] | (d[p + 1] << 8);
    }

    static uint ReadU32(byte[] d, int p)
    {
        return (uint)(d[p] | (d[p + 1] << 8) | (d[p + 2] << 16) | (d[p + 3] << 24));
    }

    static byte[] Slice(byte[] d, int off, int len)
    {
        byte[] r = new byte[len];
        Buffer.BlockCopy(d, off, r, 0, len);
        return r;
    }

    class Biff12Reader
    {
        byte[] d;
        int pos;

        public Biff12Reader(byte[] data)
        {
            d = data;
            pos = 0;
        }

        public bool Next(out int id, out int off, out int len)
        {
            id = 0;
            off = 0;
            len = 0;
            if (pos >= d.Length) return false;

            int b = d[pos];
            pos++;
            id = b & 0x7F;
            if ((b & 0x80) != 0)
            {
                if (pos >= d.Length) return false;
                int b2 = d[pos];
                pos++;
                id |= (b2 & 0x7F) << 7;
            }

            if (pos >= d.Length) return false;
            b = d[pos];
            pos++;
            len = b & 0x7F;
            for (int i = 1; i < 4; i++)
            {
                if ((b & 0x80) == 0) break;
                if (pos >= d.Length) return false;
                b = d[pos];
                pos++;
                len |= (b & 0x7F) << (7 * i);
            }

            off = pos;
            pos += len;
            return pos <= d.Length;
        }
    }
}
