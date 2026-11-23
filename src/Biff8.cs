using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

class XlsSheet
{
    public SheetInfo Info;
    public int Offset;
    public int R1 = -1, C1 = -1, R2 = -1, C2 = -1;
    public Dictionary<long, byte[]> Formulas = new Dictionary<long, byte[]>();
}

class Xti
{
    public int SupBook;
    public int First;
    public int Last;
}

class SupBook
{
    public bool IsInternal;
    public int SheetCount;
}

class NameRaw
{
    public DefinedName Dn;
    public byte[] Tokens;
}

// BIFF8 (.xls, Excel 97-2003) reader. BIFF5/95 is intentionally unsupported.
class XlsBook : BookBase
{
    byte[] file;
    byte[] wb;
    Cfb cfb;
    List<SheetInfo> sheets = new List<SheetInfo>();
    List<XlsSheet> xlsSheets = new List<XlsSheet>();
    string[] sst = new string[0];
    List<int> xfFmt = new List<int>();
    Dictionary<int, string> formats = new Dictionary<int, string>();
    bool date1904;
    bool hasVba;
    List<Xti> externs = new List<Xti>();
    List<SupBook> supbooks = new List<SupBook>();
    List<NameRaw> rawNames = new List<NameRaw>();

    public XlsBook(string path)
    {
        file = File.ReadAllBytes(path);
        cfb = new Cfb(file);

        bool encrypted = false;
        try { cfb.GetStream("EncryptedPackage"); encrypted = true; }
        catch { encrypted = false; }
        if (encrypted) throw new Exception("encrypted workbook is not supported");

        try { wb = cfb.GetStream("Workbook"); }
        catch { wb = cfb.GetStream("Book"); }

        ParseGlobals();
        ResolveNames();
        ParseSheets();

        try { cfb.GetStream("dir"); hasVba = true; }
        catch { hasVba = false; }
    }

    public override List<SheetInfo> Sheets { get { return sheets; } }
    public override bool Date1904 { get { return date1904; } }
    public override bool HasVba { get { return hasVba; } }

    public override byte[] GetVbaProject()
    {
        return hasVba ? file : null;
    }

    public override void Dispose() { }

    // ------------------------------------------------------------ globals

    void ParseGlobals()
    {
        int pos = 0;
        while (pos + 4 <= wb.Length)
        {
            int id = ReadU16(wb, pos);
            int len = ReadU16(wb, pos + 2);
            int body = pos + 4;
            pos = body + len;
            if (pos > wb.Length) break;

            if (id == 0x000A) break; // EOF

            if (id == 0x0085) // BOUNDSHEET
            {
                int off = (int)ReadU32(wb, body);
                string name = ReadShortUnicode(wb, body + 6);
                SheetInfo sh = new SheetInfo(name, null);
                sh.Offset = off;
                sheets.Add(sh);
                continue;
            }

            if (id == 0x00FC) // SST + CONTINUE
            {
                List<byte[]> parts = new List<byte[]>();
                parts.Add(Slice(wb, body, len));
                int p2 = pos;
                while (p2 + 4 <= wb.Length && ReadU16(wb, p2) == 0x003C)
                {
                    int l2 = ReadU16(wb, p2 + 2);
                    parts.Add(Slice(wb, p2 + 4, l2));
                    p2 += 4 + l2;
                }
                sst = ParseSst(parts);
                pos = p2;
                continue;
            }

            if (id == 0x00E0) // XF
            {
                xfFmt.Add(ReadU16(wb, body + 2));
                continue;
            }

            if (id == 0x041E) // FORMAT
            {
                int ifmt = ReadU16(wb, body);
                formats[ifmt] = ReadUnicode(wb, body + 2);
                continue;
            }

            if (id == 0x0018) // NAME
            {
                ParseName(body);
                continue;
            }

            if (id == 0x0017) // EXTERNSHEET
            {
                int n = ReadU16(wb, body);
                for (int i = 0; i < n; i++)
                {
                    int o = body + 2 + i * 6;
                    Xti x = new Xti();
                    x.SupBook = ReadU16(wb, o);
                    x.First = ReadU16(wb, o + 2);
                    x.Last = ReadU16(wb, o + 4);
                    externs.Add(x);
                }
                continue;
            }

            if (id == 0x01AE) // SUPBOOK
            {
                int numSheets = ReadU16(wb, body);
                bool internalSup = (body + 3 < wb.Length && wb[body + 2] == 0x01 && wb[body + 3] == 0x04);
                SupBook sb = new SupBook();
                sb.IsInternal = internalSup;
                sb.SheetCount = numSheets;
                supbooks.Add(sb);
                continue;
            }

            if (id == 0x0022) // DATEMODE
            {
                date1904 = wb[body] != 0;
                continue;
            }

            if (id == 0x002F) // FILEPASS
                throw new Exception("encrypted .xls is not supported");

            if (id == 0x0042) // CODEPAGE (BIFF8 strings are Unicode)
                continue;
        }

        for (int i = 0; i < sheets.Count; i++) allSheetNames.Add(sheets[i].Name);
    }

    void ParseName(int body)
    {
        int p = body;
        int grbit = ReadU16(wb, p); p += 2;
        p += 1; // chKey
        int cch = wb[p]; p += 1;
        int cce = ReadU16(wb, p); p += 2;
        p += 2; // ixals
        int itab = ReadU16(wb, p); p += 2;
        p += 4; // cchCustMenu..cchStatusbar

        int np;
        string name = ReadUnicodeNoCch(wb, p, cch, out np);
        byte[] tokens = Slice(wb, np, cce);

        DefinedName dn = new DefinedName();
        dn.Name = name;
        dn.LocalSheetId = (itab == 0) ? -1 : (itab - 1);
        dn.Hidden = (grbit & 0x01) != 0;
        dn.Ref = "";

        NameRaw raw = new NameRaw();
        raw.Dn = dn;
        raw.Tokens = tokens;
        rawNames.Add(raw);
        definedNames.Add(dn);
    }

    void ResolveNames()
    {
        string[] names = allSheetNames.ToArray();
        for (int i = 0; i < rawNames.Count; i++)
        {
            NameRaw raw = rawNames[i];
            FormulaContext ctx = new FormulaContext();
            ctx.SheetNames = names;
            ctx.Names = definedNames;
            ctx.ExternSheet = ExternSheet;
            ctx.RelDelta = true;
            ctx.BaseRow = 0;
            ctx.BaseCol = 0;
            raw.Dn.Ref = Ptg.Decompile(raw.Tokens, raw.Tokens.Length, ctx);
        }
    }

    string ExternSheet(int ixti)
    {
        if (ixti < 0 || ixti >= externs.Count) return null;
        Xti x = externs[ixti];
        if (x.SupBook < 0 || x.SupBook >= supbooks.Count) return null;
        if (!supbooks[x.SupBook].IsInternal) return null;

        int first = x.First;
        int last = x.Last;
        if (first == 0xFFFE || first == 0xFFFF) return null;
        if (first < 0 || first >= allSheetNames.Count) return null;
        if (last < first || last >= allSheetNames.Count) last = first;

        if (first == last) return allSheetNames[first];
        return allSheetNames[first] + ":" + allSheetNames[last];
    }

    // ------------------------------------------------------------ sheets

    void ParseSheets()
    {
        for (int i = 0; i < sheets.Count; i++)
        {
            SheetInfo sh = sheets[i];
            XlsSheet xs = new XlsSheet();
            xs.Info = sh;
            xs.Offset = sh.Offset;

            int pos = sh.Offset;
            while (pos + 4 <= wb.Length)
            {
                int id = ReadU16(wb, pos);
                int len = ReadU16(wb, pos + 2);
                int body = pos + 4;
                pos = body + len;
                if (pos > wb.Length) break;
                if (id == 0x000A) break;

                if (id == 0x0200) // DIMENSIONS
                {
                    xs.R1 = (int)ReadU32(wb, body);
                    xs.R2 = (int)ReadU32(wb, body + 4) - 1;
                    xs.C1 = ReadU16(wb, body + 8);
                    xs.C2 = ReadU16(wb, body + 10) - 1;
                    continue;
                }

                if (id == 0x0006) // FORMULA
                {
                    int row = ReadU16(wb, body);
                    int col = ReadU16(wb, body + 2);
                    int cce = ReadU16(wb, body + 20);
                    byte[] tokens = Slice(wb, body + 22, cce);
                    long key = ((long)row << 16) | (uint)col;
                    xs.Formulas[key] = tokens;
                    continue;
                }

                if (id == 0x00BC) // SHRFMLA
                {
                    int row1 = ReadU16(wb, body);
                    int col1 = wb[body + 4];
                    int tokslen = ReadU16(wb, body + 8);
                    byte[] tokens = Slice(wb, body + 10, tokslen);
                    long key = ((long)row1 << 16) | (uint)col1;
                    xs.Formulas[key] = tokens;
                    continue;
                }

                if (id == 0x0221) // ARRAY
                {
                    int row1 = ReadU16(wb, body);
                    int col1 = wb[body + 4];
                    int tokslen = ReadU16(wb, body + 12);
                    byte[] tokens = Slice(wb, body + 14, tokslen);
                    long key = ((long)row1 << 16) | (uint)col1;
                    xs.Formulas[key] = tokens;
                    continue;
                }
            }

            xlsSheets.Add(xs);
        }
    }

    public override string GetDimension(SheetInfo sh)
    {
        XlsSheet xs = FindSheet(sh);
        if (xs == null || xs.R1 < 0 || xs.R2 < xs.R1) return null;
        return A1.IndexToCol(xs.C1 + 1) + (xs.R1 + 1).ToString(CultureInfo.InvariantCulture)
            + ":" + A1.IndexToCol(xs.C2 + 1) + (xs.R2 + 1).ToString(CultureInfo.InvariantCulture);
    }

    XlsSheet FindSheet(SheetInfo sh)
    {
        for (int i = 0; i < xlsSheets.Count; i++)
            if (xlsSheets[i].Info == sh) return xlsSheets[i];
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
        XlsSheet xs = FindSheet(sh);
        if (xs == null) return;

        FormulaContext ctx = new FormulaContext();
        ctx.SheetNames = allSheetNames.ToArray();
        ctx.Names = definedNames;
        ctx.ExternSheet = ExternSheet;
        ctx.SharedFormula = delegate (int row, int col)
        {
            long key = ((long)row << 16) | (uint)col;
            byte[] t;
            if (xs.Formulas.TryGetValue(key, out t)) return t;
            return null;
        };

        bool hasPending = false;
        int pRow = 0, pCol = 0, pXf = 0;
        string pFormula = null;

        int pos = xs.Offset;
        while (pos + 4 <= wb.Length)
        {
            int id = ReadU16(wb, pos);
            int len = ReadU16(wb, pos + 2);
            int body = pos + 4;
            pos = body + len;
            if (pos > wb.Length) break;
            if (id == 0x000A) break;

            if (id == 0x00FD) // LABELSST
            {
                int row = ReadU16(wb, body);
                int col = ReadU16(wb, body + 2);
                int xf = ReadU16(wb, body + 4);
                int isst = (int)ReadU32(wb, body + 6);
                if (InRange(row + 1, col + 1, r1, c1, r2, c2))
                {
                    string v = (isst >= 0 && isst < sst.Length) ? sst[isst] : "";
                    sink(sh.Name, row + 1, col + 1, v, "s", null, xf);
                }
                continue;
            }

            if (id == 0x0203) // NUMBER
            {
                int row = ReadU16(wb, body);
                int col = ReadU16(wb, body + 2);
                int xf = ReadU16(wb, body + 4);
                double v = BitConverter.ToDouble(wb, body + 6);
                if (InRange(row + 1, col + 1, r1, c1, r2, c2))
                    sink(sh.Name, row + 1, col + 1, Num(v), "n", null, xf);
                continue;
            }

            if (id == 0x027E) // RK
            {
                int row = ReadU16(wb, body);
                int col = ReadU16(wb, body + 2);
                int xf = ReadU16(wb, body + 4);
                int rk = (int)ReadU32(wb, body + 6);
                if (InRange(row + 1, col + 1, r1, c1, r2, c2))
                    sink(sh.Name, row + 1, col + 1, Num(Rk(rk)), "n", null, xf);
                continue;
            }

            if (id == 0x00BD) // MULRK
            {
                int row = ReadU16(wb, body);
                int colFirst = ReadU16(wb, body + 2);
                int n = (len - 6) / 6;
                for (int i = 0; i < n; i++)
                {
                    int o = body + 4 + i * 6;
                    int xf = ReadU16(wb, o);
                    int rk = (int)ReadU32(wb, o + 2);
                    int col = colFirst + i;
                    if (InRange(row + 1, col + 1, r1, c1, r2, c2))
                        sink(sh.Name, row + 1, col + 1, Num(Rk(rk)), "n", null, xf);
                }
                continue;
            }

            if (id == 0x0205) // BOOLERR
            {
                int row = ReadU16(wb, body);
                int col = ReadU16(wb, body + 2);
                int xf = ReadU16(wb, body + 4);
                int val = wb[body + 6];
                bool isErr = wb[body + 7] != 0;
                if (InRange(row + 1, col + 1, r1, c1, r2, c2))
                {
                    if (isErr) sink(sh.Name, row + 1, col + 1, ErrText(val), "e", null, xf);
                    else sink(sh.Name, row + 1, col + 1, val != 0 ? "true" : "false", "b", null, xf);
                }
                continue;
            }

            if (id == 0x0204) // LABEL
            {
                int row = ReadU16(wb, body);
                int col = ReadU16(wb, body + 2);
                int xf = ReadU16(wb, body + 4);
                string v = ReadUnicode(wb, body + 6);
                if (InRange(row + 1, col + 1, r1, c1, r2, c2))
                    sink(sh.Name, row + 1, col + 1, v, "s", null, xf);
                continue;
            }

            if (id == 0x0201) // BLANK
            {
                int row = ReadU16(wb, body);
                int col = ReadU16(wb, body + 2);
                int xf = ReadU16(wb, body + 4);
                if (InRange(row + 1, col + 1, r1, c1, r2, c2))
                    sink(sh.Name, row + 1, col + 1, null, null, null, xf);
                continue;
            }

            if (id == 0x00BE) // MULBLANK
            {
                int row = ReadU16(wb, body);
                int colFirst = ReadU16(wb, body + 2);
                int n = (len - 6) / 2;
                for (int i = 0; i < n; i++)
                {
                    int col = colFirst + i;
                    int xf = ReadU16(wb, body + 4 + i * 2);
                    if (InRange(row + 1, col + 1, r1, c1, r2, c2))
                        sink(sh.Name, row + 1, col + 1, null, null, null, xf);
                }
                continue;
            }

            if (id == 0x0006) // FORMULA
            {
                int row = ReadU16(wb, body);
                int col = ReadU16(wb, body + 2);
                int xf = ReadU16(wb, body + 4);
                byte[] res = Slice(wb, body + 6, 8);
                int cce = ReadU16(wb, body + 20);
                byte[] tokens = Slice(wb, body + 22, cce);

                ctx.RelDelta = false;
                ctx.BaseRow = 0;
                ctx.BaseCol = 0;
                ctx.CellRow = row;
                ctx.CellCol = col;
                string formula = Ptg.Decompile(tokens, tokens.Length, ctx);

                if (!InRange(row + 1, col + 1, r1, c1, r2, c2)) continue;

                if (res[6] == 0xFF && res[7] == 0xFF)
                {
                    int first = res[0];
                    if (first == 0)
                    {
                        hasPending = true;
                        pRow = row; pCol = col; pXf = xf; pFormula = formula;
                    }
                    else if (first == 1)
                    {
                        sink(sh.Name, row + 1, col + 1, res[2] != 0 ? "true" : "false", "b", formula, xf);
                    }
                    else if (first == 2)
                    {
                        sink(sh.Name, row + 1, col + 1, ErrText(res[2]), "e", formula, xf);
                    }
                    else if (first == 3)
                    {
                        sink(sh.Name, row + 1, col + 1, "", "s", formula, xf);
                    }
                }
                else
                {
                    double v = BitConverter.ToDouble(res, 0);
                    sink(sh.Name, row + 1, col + 1, Num(v), "n", formula, xf);
                }
                continue;
            }

            if (id == 0x0207) // STRING (formula string result)
            {
                if (hasPending)
                {
                    string v = ReadUnicode(wb, body);
                    sink(sh.Name, pRow + 1, pCol + 1, v, "s", pFormula, pXf);
                    hasPending = false;
                    pFormula = null;
                }
                continue;
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
            default: return "#ERR" + code.ToString(CultureInfo.InvariantCulture);
        }
    }

    // ------------------------------------------------------------ strings

    static string ReadShortUnicode(byte[] d, int pos)
    {
        int cch = d[pos];
        int flags = d[pos + 1];
        int p = pos + 2;
        if ((flags & 0x01) != 0) return Encoding.Unicode.GetString(d, p, cch * 2);
        return Encoding.GetEncoding(28591).GetString(d, p, cch);
    }

    static string ReadUnicode(byte[] d, int pos)
    {
        int cch = ReadU16(d, pos);
        int flags = d[pos + 2];
        int p = pos + 3;
        if ((flags & 0x01) != 0) return Encoding.Unicode.GetString(d, p, cch * 2);
        return Encoding.GetEncoding(28591).GetString(d, p, cch);
    }

    static string ReadUnicodeNoCch(byte[] d, int pos, int cch, out int newPos)
    {
        int flags = d[pos];
        int p = pos + 1;
        string s;
        if ((flags & 0x01) != 0)
        {
            s = Encoding.Unicode.GetString(d, p, cch * 2);
            p += cch * 2;
        }
        else
        {
            s = Encoding.GetEncoding(28591).GetString(d, p, cch);
            p += cch;
        }
        newPos = p;
        return s;
    }

    static string[] ParseSst(List<byte[]> parts)
    {
        int nstrings = (int)ReadU32(parts[0], 4);
        string[] result = new string[nstrings];

        int di = 0;
        byte[] data = parts[0];
        int datalen = data.Length;
        int pos = 8;

        for (int si = 0; si < nstrings; si++)
        {
            if (pos + 2 > datalen) break;
            int nchars = ReadU16(data, pos);
            pos += 2;
            int options = data[pos];
            pos += 1;

            int rt = 0, ph = 0;
            if ((options & 0x08) != 0) { rt = ReadU16(data, pos); pos += 2; }
            if ((options & 0x04) != 0) { ph = (int)ReadU32(data, pos); pos += 4; }

            StringBuilder sb = new StringBuilder(nchars);
            int got = 0;
            while (true)
            {
                int need = nchars - got;
                if ((options & 0x01) != 0)
                {
                    int avail = Math.Min((datalen - pos) >> 1, need);
                    sb.Append(Encoding.Unicode.GetString(data, pos, avail * 2));
                    pos += avail * 2;
                    got += avail;
                }
                else
                {
                    int avail = Math.Min(datalen - pos, need);
                    sb.Append(Encoding.GetEncoding(28591).GetString(data, pos, avail));
                    pos += avail;
                    got += avail;
                }
                if (got == nchars) break;
                di++;
                if (di >= parts.Count) break;
                data = parts[di];
                datalen = data.Length;
                options = data[0];
                pos = 1;
            }

            if (rt > 0)
            {
                for (int r = 0; r < rt; r++)
                {
                    if (pos == datalen)
                    {
                        di++;
                        if (di >= parts.Count) break;
                        data = parts[di];
                        datalen = data.Length;
                        pos = 0;
                    }
                    pos += 4;
                }
            }

            pos += ph;
            if (pos >= datalen)
            {
                pos -= datalen;
                di++;
                if (di < parts.Count)
                {
                    data = parts[di];
                    datalen = data.Length;
                }
            }

            result[si] = sb.ToString();
        }

        return result;
    }

    // ------------------------------------------------------------ helpers

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
}
