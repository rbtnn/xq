using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

// A pragmatic subset of Excel's number-format engine, enough for the
// formats people actually put on cells: General, 0/#/?, thousands,
// percent, currency literals, scientific, dates/times and text.
// Custom formats (numFmtId >= 164) are exact; built-in date formats
// (14..22) follow the en-US table, which is the one Excel stores as ids.
class DateToken
{
    public char Kind;     // L literal, Y M D H S A F E N
    public int Len;
    public string Lit;
    public bool Minute;
}

static class NumFmt
{
    static readonly string[] MonthAbbr = { "Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec" };
    static readonly string[] MonthFull = { "January","February","March","April","May","June","July","August","September","October","November","December" };
    static readonly string[] DayAbbr   = { "Sun","Mon","Tue","Wed","Thu","Fri","Sat" };
    static readonly string[] DayFull   = { "Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday" };
    static readonly string[] JaDayAbbr = { "日","月","火","水","木","金","土" };
    static readonly string[] JaDayFull = { "日曜日","月曜日","火曜日","水曜日","木曜日","金曜日","土曜日" };

    // ------------------------------------------------------------ entry

    public static string Format(double v, string code, bool date1904)
    {
        if (string.IsNullOrEmpty(code) || string.Equals(code, "General", StringComparison.OrdinalIgnoreCase))
            return General(v);

        List<string> sections = SplitSections(code);
        int idx;
        string sec = PickSection(sections, v, out idx);
        if (sec == null) return General(v);
        if (sec.Trim().Length == 0) return "";
        if (IsDateFormat(sec)) return FormatDate(v, sec, date1904);
        return FormatNumber(v, sec, idx, sections.Count);
    }

    public static string FormatText(string text, string code)
    {
        if (string.IsNullOrEmpty(code)) return text;
        List<string> sections = SplitSections(code);
        string sec = sections.Count >= 4 ? sections[3] : sections[0];
        if (!HasAt(sec)) return text;
        if (sec.Trim().Length == 0) return "";
        return EmitText(sec, text);
    }

    static bool HasAt(string s)
    {
        bool inQuote = false;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == '"') { inQuote = !inQuote; continue; }
            if (inQuote) continue;
            if (c == '\\') { i++; continue; }
            if (c == '[') { int j = s.IndexOf(']', i); i = j >= 0 ? j : s.Length; continue; }
            if (c == '@') return true;
        }
        return false;
    }

    // ------------------------------------------------------------ sections

    static List<string> SplitSections(string code)
    {
        List<string> parts = new List<string>();
        StringBuilder cur = new StringBuilder();
        bool inQuote = false;
        for (int i = 0; i < code.Length; i++)
        {
            char c = code[i];
            if (c == '"') { inQuote = !inQuote; cur.Append(c); continue; }
            if (inQuote) { cur.Append(c); continue; }
            if (c == '\\') { cur.Append(c); if (i + 1 < code.Length) cur.Append(code[++i]); continue; }
            if (c == '_' || c == '*') { cur.Append(c); if (i + 1 < code.Length) cur.Append(code[++i]); continue; }
            if (c == ';') { parts.Add(cur.ToString()); cur.Length = 0; continue; }
            cur.Append(c);
        }
        parts.Add(cur.ToString());
        return parts;
    }

    static string PickSection(List<string> sections, double v, out int idx)
    {
        idx = 0;
        if (sections.Count == 1) return sections[0];
        if (sections.Count == 2)
        {
            if (v < 0) { idx = 1; return sections[1]; }
            return sections[0];
        }
        if (v > 0) { idx = 0; return sections[0]; }
        if (v < 0) { idx = 1; return sections[1]; }
        idx = 2;
        return sections[2];
    }

    // ------------------------------------------------------------ detect date

    static bool IsDateFormat(string s)
    {
        bool inQuote = false;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == '"') { inQuote = !inQuote; continue; }
            if (inQuote) continue;
            if (c == '\\') { i++; continue; }
            if (c == '[')
            {
                int end = s.IndexOf(']', i);
                if (end > i)
                {
                    string inner = s.Substring(i + 1, end - i - 1).ToLowerInvariant();
                    if (inner == "h" || inner == "hh" || inner == "m" || inner == "mm" || inner == "s" || inner == "ss")
                        return true;
                    i = end;
                }
                continue;
            }
            char lc = char.ToLowerInvariant(c);
            if (lc == 'y' || lc == 'm' || lc == 'd' || lc == 'h' || lc == 's') return true;
            if (lc == 'a')
            {
                if (i + 5 <= s.Length && s.Substring(i, 5).Equals("AM/PM", StringComparison.OrdinalIgnoreCase)) return true;
                if (i + 3 <= s.Length && s.Substring(i, 3).Equals("A/P", StringComparison.OrdinalIgnoreCase)) return true;
                int j = i;
                while (j < s.Length && char.ToLowerInvariant(s[j]) == 'a') j++;
                if (j - i >= 3) return true;
                i = j - 1;
            }
        }
        return false;
    }

    // ------------------------------------------------------------ General

    static string General(double v)
    {
        if (v == 0) return "0";
        bool neg = v < 0;
        double a = Math.Abs(v);
        string s;

        if (a >= 1e11)
        {
            s = a.ToString("0.#####E+00", CultureInfo.InvariantCulture);
            int e = s.IndexOf('E');
            if (e > 0)
            {
                string m = s.Substring(0, e);
                if (m.IndexOf('.') >= 0) m = m.TrimEnd('0').TrimEnd('.');
                s = m + s.Substring(e);
            }
        }
        else
        {
            int exp = (int)Math.Floor(Math.Log10(a));
            int decimals = 9 - exp;
            if (decimals < 0) decimals = 0;
            if (decimals > 15) decimals = 15;

            double r = Math.Round(a, decimals, MidpointRounding.AwayFromZero);
            s = r.ToString("F" + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            if (s.IndexOf('.') >= 0) s = s.TrimEnd('0').TrimEnd('.');

            // keep it within a jq-friendly 11 characters, like Excel's default width
            while (s.Length > 11 && decimals > 0)
            {
                decimals--;
                r = Math.Round(a, decimals, MidpointRounding.AwayFromZero);
                s = r.ToString("F" + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
                if (s.IndexOf('.') >= 0) s = s.TrimEnd('0').TrimEnd('.');
            }
        }

        return (neg ? "-" : "") + s;
    }

    // ------------------------------------------------------------ numbers

    static void AppendSeg(StringBuilder prefix, StringBuilder suffix, bool inPattern, char c)
    {
        if (inPattern) suffix.Append(c);
        else prefix.Append(c);
    }

    static string FormatNumber(double v, string sec, int sectionIndex, int sectionCount)
    {
        StringBuilder prefix = new StringBuilder();
        StringBuilder suffix = new StringBuilder();
        StringBuilder pat = new StringBuilder();
        bool inPattern = false;
        bool patternDone = false;
        bool inQuote = false;
        int percent = 0;
        bool hasSlash = false;
        bool seenE = false;

        for (int i = 0; i < sec.Length; i++)
        {
            char c = sec[i];
            if (c == '"') { inQuote = !inQuote; continue; }
            if (inQuote) { AppendSeg(prefix, suffix, inPattern, c); continue; }
            if (c == '\\') { i++; if (i < sec.Length) AppendSeg(prefix, suffix, inPattern, sec[i]); continue; }
            if (c == '_') { AppendSeg(prefix, suffix, inPattern, ' '); i++; continue; }
            if (c == '*') { i++; continue; }
            if (c == '[')
            {
                int j = sec.IndexOf(']', i);
                string inner = j > i ? sec.Substring(i + 1, j - i - 1) : "";
                if (inner.StartsWith("$"))
                {
                    int dash = inner.IndexOf('-');
                    string sym = inner.Substring(1, dash > 1 ? dash - 1 : inner.Length - 1);
                    for (int k = 0; k < sym.Length; k++) AppendSeg(prefix, suffix, inPattern, sym[k]);
                }
                i = j >= 0 ? j : sec.Length;
                continue;
            }
            if (c == '%') { percent++; AppendSeg(prefix, suffix, inPattern, '%'); continue; }

            if (!patternDone && (c == '0' || c == '#' || c == '?' || c == '.' || c == ',' || c == 'E' || c == 'e' || c == '/'))
            {
                inPattern = true;
                pat.Append(c);
                if (c == 'E' || c == 'e') seenE = true;
                if (c == '/') hasSlash = true;
                continue;
            }
            if (!patternDone && seenE && (c == '+' || c == '-'))
            {
                pat.Append(c);
                continue;
            }
            if (inPattern) patternDone = true;
            AppendSeg(prefix, suffix, inPattern, c);
        }

        if (hasSlash)
            return General(v); // fractions are not implemented yet

        string p = pat.ToString();
        int ePos = -1;
        for (int i = 0; i < p.Length; i++) if (p[i] == 'E' || p[i] == 'e') { ePos = i; break; }
        string mant = ePos >= 0 ? p.Substring(0, ePos) : p;
        string expPat = ePos >= 0 ? p.Substring(ePos + 1) : "";

        int scale = 0;
        while (mant.Length > 0 && mant[mant.Length - 1] == ',') { scale++; mant = mant.Substring(0, mant.Length - 1); }

        int dot = mant.IndexOf('.');
        string intPat = dot >= 0 ? mant.Substring(0, dot) : mant;
        string decPat = dot >= 0 ? mant.Substring(dot + 1) : "";
        bool grouping = intPat.IndexOf(',') >= 0;

        int minInt = 0, maxInt = 0;
        for (int i = 0; i < intPat.Length; i++)
        {
            if (intPat[i] == '0') { minInt++; maxInt++; }
            else if (intPat[i] == '#' || intPat[i] == '?') maxInt++;
        }
        int decCount = 0, minDec = 0;
        for (int i = 0; i < decPat.Length; i++)
        {
            if (decPat[i] == '0') { minDec++; decCount++; }
            else if (decPat[i] == '#' || decPat[i] == '?') decCount++;
        }

        double x = v;
        bool neg = x < 0;
        if (sectionCount == 1)
        {
            if (neg) x = -x;
        }
        else
        {
            x = Math.Abs(x);
        }
        for (int i = 0; i < percent; i++) x *= 100;
        if (scale > 0) x /= Math.Pow(1000, scale);
        if (x == 0) x = 0;

        string num;
        if (ePos >= 0)
        {
            num = Scientific(x, decCount, expPat, maxInt, decPat);
        }
        else
        {
            int dc = decCount;
            if (dc > 15) dc = 15;
            double r = Math.Round(x, dc, MidpointRounding.AwayFromZero);
            string s = r.ToString("F" + dc.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            int dpos = s.IndexOf('.');
            string ip = dpos >= 0 ? s.Substring(0, dpos) : s;
            string dp = dpos >= 0 ? s.Substring(dpos + 1) : "";

            if (maxInt == 0) ip = "";
            else
            {
                if (ip == "0" && minInt == 0) ip = "";
                if (ip.Length < minInt) ip = new string('0', minInt - ip.Length) + ip;
                if (grouping) ip = Group(ip);
            }

            if (decCount > 0)
            {
                StringBuilder db = new StringBuilder();
                for (int i = 0; i < decCount && i < dp.Length; i++)
                {
                    char pc = decPat[i];
                    char dc2 = dp[i];
                    if (dc2 == '0')
                    {
                        if (pc == '#') break;
                        if (pc == '?') { db.Append(' '); continue; }
                    }
                    db.Append(dc2);
                }
                dp = db.ToString();
                num = ip + "." + dp;
            }
            else
            {
                num = ip;
            }
        }

        if (sectionCount == 1 && neg) prefix.Insert(0, '-');
        return prefix.ToString() + num + suffix.ToString();
    }

    static string Scientific(double x, int decCount, string expPat, int maxInt, string decPat)
    {
        int exp;
        if (x == 0) exp = 0;
        else exp = (int)Math.Floor(Math.Log10(x));
        if (maxInt >= 3) exp = (int)(Math.Floor(exp / 3.0) * 3);

        double mant = x == 0 ? 0 : x / Math.Pow(10, exp);
        int dc = decCount;
        if (dc > 15) dc = 15;
        mant = Math.Round(mant, dc, MidpointRounding.AwayFromZero);

        if (maxInt >= 3 && mant >= 1000) { mant /= 1000; exp += 3; }
        else if (maxInt < 3 && mant >= 10) { mant /= 10; exp++; }

        string ms = mant.ToString("F" + dc.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        if (dc > 0 && ms.IndexOf('.') >= 0)
        {
            int minDec = 0;
            for (int i = 0; i < decPat.Length; i++) if (decPat[i] == '0') minDec++;

            int dotPos = ms.IndexOf('.');
            string digits = ms.Substring(dotPos + 1);
            int keep = digits.Length;
            while (keep > minDec && digits[keep - 1] == '0') keep--;

            ms = ms.Substring(0, dotPos);
            if (keep > 0) ms += "." + digits.Substring(0, keep);
        }

        int expDigits = 0;
        for (int i = 0; i < expPat.Length; i++) if (expPat[i] >= '0' && expPat[i] <= '9') expDigits++;
        if (expDigits < 1) expDigits = 2;

        string es = Math.Abs(exp).ToString(CultureInfo.InvariantCulture).PadLeft(expDigits, '0');
        char signChar = expPat.Length > 0 && expPat[0] == '+' ? '+' : '-';
        string sign = exp < 0 ? "-" : (signChar == '+' ? "+" : "");
        return ms + "E" + sign + es;
    }

    static string Group(string digits)
    {
        if (digits.Length <= 3) return digits;
        StringBuilder sb = new StringBuilder(digits.Length + digits.Length / 3);
        int first = digits.Length % 3;
        if (first > 0) { sb.Append(digits, 0, first); sb.Append(','); }
        for (int i = first; i < digits.Length; i += 3)
        {
            sb.Append(digits, i, 3);
            if (i + 3 < digits.Length) sb.Append(',');
        }
        return sb.ToString();
    }

    // ------------------------------------------------------------ dates

    static List<DateToken> ParseDate(string s)
    {
        List<DateToken> list = new List<DateToken>();
        int i = 0;
        while (i < s.Length)
        {
            char c = s[i];

            if (c == '"')
            {
                int j = s.IndexOf('"', i + 1);
                string lit = j > i ? s.Substring(i + 1, j - i - 1) : "";
                list.Add(Lit(lit));
                i = j > i ? j + 1 : i + 1;
                continue;
            }
            if (c == '\\')
            {
                if (i + 1 < s.Length) list.Add(Lit(s[i + 1].ToString()));
                i += 2;
                continue;
            }
            if (c == '_') { list.Add(Lit(" ")); i += 2; continue; }
            if (c == '*') { i += 2; continue; }

            if (c == '[')
            {
                int j = s.IndexOf(']', i);
                if (j < 0) { list.Add(Lit(s.Substring(i))); break; }
                string inner = s.Substring(i + 1, j - i - 1);
                string low = inner.ToLowerInvariant();
                if (low == "h" || low == "hh") list.Add(Elapsed("h", low.Length));
                else if (low == "m" || low == "mm") list.Add(Elapsed("m", low.Length));
                else if (low == "s" || low == "ss") list.Add(Elapsed("s", low.Length));
                else if (inner.StartsWith("$"))
                {
                    int dash = inner.IndexOf('-');
                    string sym = inner.Substring(1, dash > 1 ? dash - 1 : inner.Length - 1);
                    if (sym.Length > 0) list.Add(Lit(sym));
                }
                i = j + 1;
                continue;
            }

            if (i + 5 <= s.Length && s.Substring(i, 5).Equals("AM/PM", StringComparison.OrdinalIgnoreCase))
            { list.Add(new DateToken { Kind = 'A', Len = 5 }); i += 5; continue; }
            if (i + 3 <= s.Length && s.Substring(i, 3).Equals("A/P", StringComparison.OrdinalIgnoreCase))
            { list.Add(new DateToken { Kind = 'A', Len = 3 }); i += 3; continue; }

            if (c == '.' && i + 1 < s.Length && s[i + 1] == '0')
            {
                int j = i + 1;
                while (j < s.Length && s[j] == '0') j++;
                list.Add(new DateToken { Kind = 'F', Len = j - (i + 1) });
                i = j;
                continue;
            }

            char lc = char.ToLowerInvariant(c);
            if (lc == 'y' || lc == 'm' || lc == 'd' || lc == 'h' || lc == 's')
            {
                int j = i;
                while (j < s.Length && char.ToLowerInvariant(s[j]) == lc) j++;
                list.Add(new DateToken { Kind = char.ToUpperInvariant(lc), Len = j - i });
                i = j;
                continue;
            }
            if (lc == 'a')
            {
                int j = i;
                while (j < s.Length && char.ToLowerInvariant(s[j]) == 'a') j++;
                int len = j - i;
                if (len >= 3) list.Add(new DateToken { Kind = 'N', Len = len });
                else list.Add(Lit(s.Substring(i, len)));
                i = j;
                continue;
            }

            list.Add(Lit(c.ToString()));
            i++;
        }
        return list;
    }

    static DateToken Lit(string s)
    {
        return new DateToken { Kind = 'L', Lit = s };
    }

    static DateToken Elapsed(string which, int len)
    {
        return new DateToken { Kind = 'E', Lit = which, Len = len };
    }

    static void ResolveMinutes(List<DateToken> tokens)
    {
        for (int i = 0; i < tokens.Count; i++)
        {
            if (tokens[i].Kind != 'M') continue;
            char prev = '\0', next = '\0';
            for (int j = i - 1; j >= 0; j--) if (tokens[j].Kind != 'L') { prev = tokens[j].Kind; break; }
            for (int j = i + 1; j < tokens.Count; j++) if (tokens[j].Kind != 'L') { next = tokens[j].Kind; break; }
            tokens[i].Minute = (prev == 'H' || prev == 'E' || next == 'S');
        }
    }

    static DateTime SerialToDate(double serial, bool date1904)
    {
        double d = serial;
        if (date1904) d += 1462.0;
        if (d >= 61) return DateTime.FromOADate(d);
        if (d >= 1) return DateTime.FromOADate(d + 1);
        if (d >= 0) return DateTime.FromOADate(d);
        return DateTime.FromOADate(0);
    }

    static string Pad2(int n)
    {
        return n.ToString("00", CultureInfo.InvariantCulture);
    }

    static string FormatDate(double v, string sec, bool date1904)
    {
        List<DateToken> tokens = ParseDate(sec);
        ResolveMinutes(tokens);
        DateTime dt = SerialToDate(v, date1904);
        double adj = v + (date1904 ? 1462.0 : 0.0);
        bool fakeFeb29 = (adj == 60.0);

        bool ampm = false;
        for (int i = 0; i < tokens.Count; i++) if (tokens[i].Kind == 'A') ampm = true;

        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < tokens.Count; i++)
        {
            DateToken t = tokens[i];
            switch (t.Kind)
            {
                case 'L':
                    sb.Append(t.Lit);
                    break;

                case 'Y':
                    {
                        int y = fakeFeb29 ? 1900 : dt.Year;
                        if (t.Len == 2) sb.Append(Pad2(y % 100));
                        else sb.Append(y.ToString("0000", CultureInfo.InvariantCulture));
                    }
                    break;

                case 'M':
                    if (t.Minute)
                    {
                        if (t.Len >= 2) sb.Append(Pad2(dt.Minute));
                        else sb.Append(dt.Minute.ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        int mo = fakeFeb29 ? 2 : dt.Month;
                        if (t.Len == 1) sb.Append(mo.ToString(CultureInfo.InvariantCulture));
                        else if (t.Len == 2) sb.Append(Pad2(mo));
                        else if (t.Len == 3) sb.Append(MonthAbbr[mo - 1]);
                        else if (t.Len == 4) sb.Append(MonthFull[mo - 1]);
                        else sb.Append(MonthFull[mo - 1].Substring(0, 1));
                    }
                    break;

                case 'D':
                    {
                        int day = fakeFeb29 ? 29 : dt.Day;
                        if (t.Len == 1) sb.Append(day.ToString(CultureInfo.InvariantCulture));
                        else if (t.Len == 2) sb.Append(Pad2(day));
                        else if (t.Len == 3) sb.Append(DayAbbr[(int)dt.DayOfWeek]);
                        else sb.Append(DayFull[(int)dt.DayOfWeek]);
                    }
                    break;

                case 'H':
                    int h = ampm ? (dt.Hour % 12 == 0 ? 12 : dt.Hour % 12) : dt.Hour;
                    if (t.Len >= 2) sb.Append(Pad2(h));
                    else sb.Append(h.ToString(CultureInfo.InvariantCulture));
                    break;

                case 'S':
                    if (t.Len >= 2) sb.Append(Pad2(dt.Second));
                    else sb.Append(dt.Second.ToString(CultureInfo.InvariantCulture));
                    break;

                case 'A':
                    if (t.Len == 5) sb.Append(dt.Hour < 12 ? "AM" : "PM");
                    else sb.Append(dt.Hour < 12 ? "A" : "P");
                    break;

                case 'F':
                    {
                        double totalSec = v * 86400.0;
                        double frac = totalSec - Math.Floor(totalSec);
                        int digits = t.Len;
                        if (digits > 7) digits = 7;
                        double scaled = frac * Math.Pow(10, digits);
                        int val = (int)Math.Round(scaled, MidpointRounding.AwayFromZero);
                        int cap = (int)Math.Pow(10, digits);
                        if (val >= cap) val = cap - 1;
                        sb.Append(val.ToString("D" + digits.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture));
                    }
                    break;

                case 'E':
                    {
                        double total;
                        if (t.Lit == "h") total = Math.Floor(v * 24.0 + 1e-9);
                        else if (t.Lit == "m") total = Math.Floor(v * 1440.0 + 1e-9);
                        else total = Math.Floor(v * 86400.0 + 1e-9);
                        string es = ((long)total).ToString(CultureInfo.InvariantCulture);
                        if (t.Len >= 2 && es.Length < 2) es = es.PadLeft(2, '0');
                        sb.Append(es);
                    }
                    break;

                case 'N':
                    if (t.Len >= 4) sb.Append(JaDayFull[(int)dt.DayOfWeek]);
                    else sb.Append(JaDayAbbr[(int)dt.DayOfWeek]);
                    break;
            }
        }
        return sb.ToString();
    }

    // ------------------------------------------------------------ text

    static string EmitText(string sec, string text)
    {
        StringBuilder sb = new StringBuilder();
        bool inQuote = false;
        bool usedAt = false;

        for (int i = 0; i < sec.Length; i++)
        {
            char c = sec[i];
            if (c == '"') { inQuote = !inQuote; continue; }
            if (inQuote) { sb.Append(c); continue; }
            if (c == '\\') { i++; if (i < sec.Length) sb.Append(sec[i]); continue; }
            if (c == '_') { sb.Append(' '); i++; continue; }
            if (c == '*') { i++; continue; }
            if (c == '[')
            {
                int j = sec.IndexOf(']', i);
                i = j >= 0 ? j : sec.Length;
                continue;
            }
            if (c == '@') { sb.Append(text); usedAt = true; continue; }
            sb.Append(c);
        }

        if (!usedAt) return sb.ToString();
        return sb.ToString();
    }
}
