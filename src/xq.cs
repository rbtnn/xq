using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

class LimitReachedException : Exception { }

class Program
{
    const string VERSION = "xq 0.1.0";

    static int Main(string[] args)
    {
        try { Console.OutputEncoding = new UTF8Encoding(false); }
        catch { }

        if (args.Length == 0) { Usage(); return 2; }

        string cmd = args[0];
        if (cmd == "--help" || cmd == "-h" || cmd == "help") { Usage(); return 0; }
        if (cmd == "--version" || cmd == "-V") { Console.WriteLine(VERSION); return 0; }

        string file = null;
        string expr = null;
        string sheetOpt = null;
        string rangeOpt = null;
        string moduleOpt = null;
        string nameOpt = null;
        string grepOpt = null;
        bool json = false;
        bool formulas = false;
        bool all = false;
        bool text = false;
        int limitOpt = -1;

        bool knownCommand = cmd == "sheets" || cmd == "cells" || cmd == "names"
            || cmd == "vba" || cmd == "info" || cmd == "eval";
        if (!knownCommand && !cmd.StartsWith("-"))
        {
            file = cmd;
            cmd = "eval";
        }

        for (int i = 1; i < args.Length; i++)
        {
            string a = args[i];
            if (a == "--sheet")
            {
                if (i + 1 >= args.Length) { Console.Error.WriteLine("--sheet needs a value"); return 2; }
                sheetOpt = args[++i];
            }
            else if (a == "--range")
            {
                if (i + 1 >= args.Length) { Console.Error.WriteLine("--range needs a value"); return 2; }
                rangeOpt = args[++i];
            }
            else if (a == "--module")
            {
                if (i + 1 >= args.Length) { Console.Error.WriteLine("--module needs a value"); return 2; }
                moduleOpt = args[++i];
            }
            else if (a == "--name")
            {
                if (i + 1 >= args.Length) { Console.Error.WriteLine("--name needs a value"); return 2; }
                nameOpt = args[++i];
            }
            else if (a == "--grep")
            {
                if (i + 1 >= args.Length) { Console.Error.WriteLine("--grep needs a value"); return 2; }
                grepOpt = args[++i];
            }
            else if (a == "--limit")
            {
                if (i + 1 >= args.Length) { Console.Error.WriteLine("--limit needs a value"); return 2; }
                if (!int.TryParse(args[++i], NumberStyles.Integer, CultureInfo.InvariantCulture, out limitOpt) || limitOpt < 0)
                {
                    Console.Error.WriteLine("--limit needs a non-negative integer");
                    return 2;
                }
            }
            else if (a == "-e" || a == "--expr")
            {
                if (i + 1 >= args.Length) { Console.Error.WriteLine("-e needs a value"); return 2; }
                expr = args[++i];
            }
            else if (a == "--json") json = true;
            else if (a == "--formulas") formulas = true;
            else if (a == "--all") all = true;
            else if (a == "--text") text = true;
            else if (a == "--help" || a == "-h") { Usage(); return 0; }
            else if (a.Length > 2 && a[0] == '-' && a[1] == '-')
            {
                Console.Error.WriteLine("unknown option: " + a);
                return 2;
            }
            else if (file == null) file = a;
            else { Console.Error.WriteLine("unexpected argument: " + a); return 2; }
        }

        if (file == null) { Console.Error.WriteLine("missing file"); return 2; }

        try
        {
            if (cmd == "sheets") return CmdSheets(file);
            if (cmd == "cells") return CmdCells(file, sheetOpt, rangeOpt, json, formulas, all, text, limitOpt);
            if (cmd == "eval") return CmdEval(file, expr, sheetOpt, json, text);
            if (cmd == "names") return CmdNames(file, nameOpt, sheetOpt, grepOpt, json);
            if (cmd == "vba") return CmdVba(file, moduleOpt);
            if (cmd == "info") return CmdInfo(file);

            Console.Error.WriteLine("unknown command: " + cmd);
            Usage();
            return 2;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("xq: " + ex.Message);
            return 1;
        }
    }

    // ------------------------------------------------------------ commands

    static int CmdSheets(string file)
    {
        IBook x = Book.Open(file);
        StreamWriter sw = Out();
        StringBuilder sb = new StringBuilder();
        sb.Append('[');
        for (int i = 0; i < x.Sheets.Count; i++)
        {
            if (i > 0) sb.Append(',');
            JsonString(sb, x.Sheets[i].Name);
        }
        sb.Append(']');
        sw.WriteLine(sb.ToString());
        sw.Flush();
        x.Dispose();
        return 0;
    }

    static int CmdCells(string file, string sheetOpt, string rangeOpt, bool json, bool formulas, bool all, bool text, int limit)
    {
        IBook x = Book.Open(file);

        int r1 = 1, c1 = 1, r2 = int.MaxValue, c2 = int.MaxValue;
        if (rangeOpt != null)
        {
            if (!A1.TryParseRange(rangeOpt, out r1, out c1, out r2, out c2))
            {
                string targetSheet;
                if (!x.TryResolveNameRange(rangeOpt, sheetOpt, out targetSheet, out r1, out c1, out r2, out c2))
                {
                    if (x.FindName(rangeOpt, sheetOpt) != null)
                        Console.Error.WriteLine("defined name is not a single rectangular range: " + rangeOpt);
                    else
                        Console.Error.WriteLine("bad --range or unknown name: " + rangeOpt);
                    return 2;
                }
                if (sheetOpt == null && targetSheet != null) sheetOpt = targetSheet;
            }
        }
        StreamWriter sw = Out();
        StringBuilder sb = new StringBuilder(192);
        StringBuilder outBuf = new StringBuilder(65536);

        bool first = true;
        int emitted = 0;
        if (json) sw.Write('[');

        CellSink sink = delegate (string sn, int row, int col, string val, string type, string f, int styleIndex)
        {
            if (val == null && !all) return;
            if (limit >= 0 && emitted >= limit) throw new LimitReachedException();

            sb.Length = 0;
            sb.Append("{\"sheet\":");
            JsonString(sb, sn);
            sb.Append(",\"ref\":\"");
            AppendCol(sb, col);
            sb.Append(row);
            sb.Append("\",\"r\":");
            sb.Append(row);
            sb.Append(",\"c\":");
            sb.Append(col);
            sb.Append(",\"v\":");

            if (val == null) sb.Append("null");
            else if (type == "n" || type == "b") sb.Append(val);
            else JsonString(sb, val);

            sb.Append(",\"t\":");
            if (type == null) sb.Append("null");
            else JsonString(sb, type);

            if (formulas && f != null && f.Length > 0)
            {
                sb.Append(",\"f\":");
                JsonString(sb, f);
            }

            if (text)
            {
                sb.Append(",\"text\":");
                string tv = CellText(x, styleIndex, val, type);
                if (tv == null) sb.Append("null");
                else JsonString(sb, tv);
            }
            sb.Append('}');

            if (json)
            {
                if (!first) outBuf.Append(',');
                first = false;
            }
            outBuf.Append(sb);
            if (!json) outBuf.Append('\n');

            emitted++;
            if (limit >= 0 && emitted >= limit) throw new LimitReachedException();

            if (outBuf.Length >= 65536)
            {
                sw.Write(outBuf.ToString());
                outBuf.Length = 0;
            }
        };

        try
        {
            for (int i = 0; i < x.Sheets.Count; i++)
            {
                SheetInfo sh = x.Sheets[i];
                if (sheetOpt != null && !string.Equals(sh.Name, sheetOpt, StringComparison.Ordinal)) continue;

                int rr1 = r1;
                int cc1 = c1;
                int rr2 = r2;
                int cc2 = c2;
                ClampRange(x, sh, ref rr1, ref cc1, ref rr2, ref cc2);
                x.ReadCells(sh, rr1, cc1, rr2, cc2, sink);
            }
        }
        catch (LimitReachedException)
        {
        }

        if (outBuf.Length > 0)
        {
            sw.Write(outBuf.ToString());
            outBuf.Length = 0;
        }

        if (json)
        {
            sw.Write(']');
            sw.Write('\n');
        }
        sw.Flush();
        x.Dispose();
        return 0;
    }

    static int CmdVba(string file, string moduleOpt)
    {
        IBook x = Book.Open(file);
        byte[] bin = x.GetVbaProject();
        StreamWriter sw = Out();
        StringBuilder sb = new StringBuilder(4096);

        if (bin == null)
        {
            sb.Append("{\"codePage\":0,\"modules\":[]}");
            sw.WriteLine(sb.ToString());
            sw.Flush();
            x.Dispose();
            return 0;
        }

        int codePage;
        List<VbaModule> modules = Ovba.ReadProject(bin, out codePage);

        sb.Append("{\"codePage\":").Append(codePage.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"modules\":[");

        bool first = true;
        for (int i = 0; i < modules.Count; i++)
        {
            VbaModule m = modules[i];
            if (moduleOpt != null
                && !string.Equals(m.Name, moduleOpt, StringComparison.Ordinal)
                && !string.Equals(m.StreamName, moduleOpt, StringComparison.Ordinal)) continue;

            if (!first) sb.Append(',');
            first = false;

            sb.Append("{\"name\":");
            JsonString(sb, m.Name);
            sb.Append(",\"stream\":");
            JsonString(sb, m.StreamName);
            sb.Append(",\"kind\":");
            JsonString(sb, m.Kind == 0 ? "standard" : "class");
            sb.Append(",\"offset\":").Append(m.TextOffset.ToString(CultureInfo.InvariantCulture));
            sb.Append(",\"readonly\":").Append(m.IsReadOnly ? "true" : "false");
            sb.Append(",\"private\":").Append(m.IsPrivate ? "true" : "false");
            sb.Append(",\"code\":");
            JsonString(sb, m.Code);
            sb.Append('}');
        }

        sb.Append("]}");
        sw.WriteLine(sb.ToString());
        sw.Flush();
        x.Dispose();
        return 0;
    }

    static int CmdEval(string file, string expr, string sheetOpt, bool json, bool text)
    {
        if (expr == null)
        {
            Console.Error.WriteLine("missing -e expression");
            return 2;
        }

        IBook x = Book.Open(file);
        try
        {
            EvalContext ctx = new EvalContext(x, sheetOpt);
            EvalValue v;
            try
            {
                v = ctx.Evaluate(expr);
            }
            catch (EvalParseException ex)
            {
                Console.Error.WriteLine("xq: " + ex.Message);
                return 2;
            }

            StreamWriter sw = Out();
            if (json)
            {
                StringBuilder sb = new StringBuilder(128);
                sb.Append("{\"v\":");
                EvalJson.AppendValue(sb, v);
                sb.Append(",\"t\":\"");
                sb.Append(v.TypeCode());
                sb.Append('"');
                if (text)
                {
                    sb.Append(",\"text\":");
                    EvalJson.AppendString(sb, v.ToDisplayString());
                }
                sb.Append('}');
                sw.WriteLine(sb.ToString());
            }
            else
            {
                sw.WriteLine(v.ToDisplayString());
            }
            sw.Flush();
            return 0;
        }
        finally
        {
            x.Dispose();
        }
    }

    static int CmdNames(string file, string nameOpt, string sheetOpt, string grepOpt, bool json)
    {
        IBook x = Book.Open(file);
        StreamWriter sw = Out();
        StringBuilder sb = new StringBuilder(256);
        bool first = true;
        if (json) sw.Write('[');

        for (int i = 0; i < x.DefinedNames.Count; i++)
        {
            DefinedName dn = x.DefinedNames[i];

            if (nameOpt != null && !string.Equals(dn.Name, nameOpt, StringComparison.OrdinalIgnoreCase)) continue;

            if (sheetOpt != null && dn.LocalSheetId >= 0
                && !string.Equals(x.SheetName(dn.LocalSheetId), sheetOpt, StringComparison.Ordinal))
                continue;

            if (grepOpt != null
                && !Contains(dn.Name, grepOpt)
                && !Contains(dn.Ref, grepOpt)
                && !Contains(dn.Comment, grepOpt))
                continue;

            string sn = dn.LocalSheetId >= 0 ? x.SheetName(dn.LocalSheetId) : null;

            sb.Length = 0;
            sb.Append("{\"name\":");
            JsonString(sb, dn.Name);
            sb.Append(",\"scope\":");
            JsonString(sb, dn.LocalSheetId >= 0 ? "sheet" : "workbook");
            sb.Append(",\"sheet\":");
            if (sn == null) sb.Append("null"); else JsonString(sb, sn);
            sb.Append(",\"sheetId\":");
            if (dn.LocalSheetId >= 0) sb.Append(dn.LocalSheetId.ToString(CultureInfo.InvariantCulture));
            else sb.Append("null");
            sb.Append(",\"hidden\":").Append(dn.Hidden ? "true" : "false");
            sb.Append(",\"builtin\":").Append(
                dn.Name != null && dn.Name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase) ? "true" : "false");
            sb.Append(",\"ref\":");
            if (dn.Ref == null) sb.Append("null"); else JsonString(sb, dn.Ref);
            sb.Append(",\"comment\":");
            if (dn.Comment == null) sb.Append("null"); else JsonString(sb, dn.Comment);
            sb.Append('}');

            if (json)
            {
                if (!first) sw.Write(',');
                first = false;
                sw.Write(sb.ToString());
            }
            else
            {
                sw.Write(sb.ToString());
                sw.Write('\n');
            }
        }

        if (json)
        {
            sw.Write(']');
            sw.Write('\n');
        }
        sw.Flush();
        x.Dispose();
        return 0;
    }

    static bool Contains(string s, string needle)
    {
        return s != null && s.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    static int CmdInfo(string file)
    {
        FileInfo fi = new FileInfo(file);
        IBook x = Book.Open(file);
        StreamWriter sw = Out();
        StringBuilder sb = new StringBuilder(1024);

        sb.Append("{\"file\":");
        JsonString(sb, fi.Name);
        sb.Append(",\"size\":").Append(fi.Length.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"sheets\":[");

        for (int i = 0; i < x.Sheets.Count; i++)
        {
            if (i > 0) sb.Append(',');
            SheetInfo sh = x.Sheets[i];
            sb.Append("{\"name\":");
            JsonString(sb, sh.Name);
            sb.Append(",\"dimension\":");
            string dim = x.GetDimension(sh);
            if (dim == null) sb.Append("null"); else JsonString(sb, dim);
            sb.Append('}');
        }

        sb.Append("],\"hasVba\":").Append(x.HasVba ? "true" : "false");

        int modules = 0;
        int codePage = 0;
        if (x.HasVba)
        {
            byte[] bin = x.GetVbaProject();
            int cp;
            List<VbaModule> mods = Ovba.ReadProject(bin, out cp);
            modules = mods.Count;
            codePage = cp;
        }
        sb.Append(",\"vbaModules\":").Append(modules.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"vbaCodePage\":").Append(codePage.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"definedNames\":").Append(x.DefinedNames.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append('}');

        sw.WriteLine(sb.ToString());
        sw.Flush();
        x.Dispose();
        return 0;
    }

    // ------------------------------------------------------------ helpers

    static void ClampRange(IBook x, SheetInfo sh, ref int r1, ref int c1, ref int r2, ref int c2)
    {
        if (r1 > 0 && r2 > 0 && c1 > 0 && c2 > 0) return;

        int dr1, dc1, dr2, dc2;
        string dim = x.GetDimension(sh);
        if (!string.IsNullOrEmpty(dim) && A1.TryParseRange(dim, out dr1, out dc1, out dr2, out dc2))
        {
            if (r1 <= 0) r1 = dr1;
            if (r2 <= 0) r2 = dr2;
            if (c1 <= 0) c1 = dc1;
            if (c2 <= 0) c2 = dc2;
        }
        if (r1 <= 0) r1 = 1;
        if (r2 <= 0) r2 = 1;
        if (c1 <= 0) c1 = 1;
        if (c2 <= 0) c2 = 1;
    }

    static string CellText(IBook x, int styleIndex, string val, string type)
    {
        if (val == null) return null;
        string code = x.GetFormatCode(styleIndex);
        if (type == "s") return NumFmt.FormatText(val, code);
        if (type == "n")
        {
            double d;
            if (double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                return NumFmt.Format(d, code, x.Date1904);
            return val;
        }
        if (type == "b") return val == "true" ? "TRUE" : "FALSE";
        return val;
    }

    static void AppendCol(StringBuilder sb, int n)
    {
        char[] buf = new char[8];
        int i = buf.Length;
        while (n > 0)
        {
            n--;
            buf[--i] = (char)('A' + (n % 26));
            n /= 26;
        }
        sb.Append(buf, i, buf.Length - i);
    }

    static StreamWriter Out()
    {
        StreamWriter sw = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false));
        sw.AutoFlush = false;
        return sw;
    }

    static void JsonString(StringBuilder sb, string s)
    {
        sb.Append('"');
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == '"') sb.Append("\\\"");
            else if (c == '\\') sb.Append("\\\\");
            else if (c == '\n') sb.Append("\\n");
            else if (c == '\r') sb.Append("\\r");
            else if (c == '\t') sb.Append("\\t");
            else if (c == '\b') sb.Append("\\b");
            else if (c == '\f') sb.Append("\\f");
            else if (c < 0x20)
            {
                sb.Append("\\u");
                sb.Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
            }
            else
            {
                sb.Append(c);
            }
        }
        sb.Append('"');
    }

    static void Usage()
    {
        Console.Error.WriteLine(VERSION);
        Console.Error.WriteLine("usage: xq <command> <file.xlsx|file.xlsm|file.xlsb|file.xls> [options]");
        Console.Error.WriteLine();
        Console.Error.WriteLine("  sheets <file>                 sheet names as a JSON array");
        Console.Error.WriteLine("  cells  <file>                 cells as NDJSON (one object per cell)");
        Console.Error.WriteLine("         --sheet NAME           only this sheet");
        Console.Error.WriteLine("         --range A1:C10|NAME    only this rectangle (or a defined name)");
        Console.Error.WriteLine("         --json                 wrap cells in a JSON array");
        Console.Error.WriteLine("         --formulas             include \"f\" (formula text)");
        Console.Error.WriteLine("         --all                  include empty cells (\"v\":null)");
        Console.Error.WriteLine("         --text                 add \"text\" (Excel-style display string)");
        Console.Error.WriteLine("         --limit N              stop after N emitted cells");
        Console.Error.WriteLine("  eval   <file> -e EXPR        evaluate an Excel expression (e.g. -e \"=A1+B1\")");
        Console.Error.WriteLine("  <file> -e EXPR               same as eval (xq book.xlsx -e \"=A1\")");
        Console.Error.WriteLine("  names  <file>                 defined names (named ranges) as NDJSON");
        Console.Error.WriteLine("         --name NAME            exact name");
        Console.Error.WriteLine("         --grep TEXT            substring in name / ref / comment");
        Console.Error.WriteLine("         --sheet NAME           local names for this sheet (+ workbook scope)");
        Console.Error.WriteLine("  vba    <file.xlsm|xlsb|xls>   VBA modules with source code as JSON");
        Console.Error.WriteLine("         --module NAME          only this module");
        Console.Error.WriteLine("  info   <file>                 sheets, dimensions, VBA summary");
        Console.Error.WriteLine();
        Console.Error.WriteLine("examples:");
        Console.Error.WriteLine("  xq cells book.xlsx --range A1:C10 | jq -c 'select(.v==\"りんご\")'");
        Console.Error.WriteLine("  xq vba book.xlsm | jq -r '.modules[].code'");
    }
}
