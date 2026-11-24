using System;
using System.Globalization;

static class A1
{
    static bool IsAsciiLetter(char c)
    {
        return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
    }

    public static int ColToIndex(string letters)
    {
        int n = 0;
        for (int i = 0; i < letters.Length; i++)
        {
            char c = char.ToUpperInvariant(letters[i]);
            if (c < 'A' || c > 'Z') throw new FormatException("bad column: " + letters);
            n = n * 26 + (c - 'A' + 1);
        }
        return n;
    }

    public static string IndexToCol(int n)
    {
        char[] buf = new char[8];
        int i = buf.Length;
        while (n > 0)
        {
            n--;
            buf[--i] = (char)('A' + (n % 26));
            n /= 26;
        }
        return new string(buf, i, buf.Length - i);
    }

    public static bool TryParseRef(string s, out int row, out int col)
    {
        row = 0;
        col = 0;
        if (string.IsNullOrEmpty(s)) return false;

        int i = 0;
        if (s[i] == '$') i++;

        int start = i;
        while (i < s.Length && IsAsciiLetter(s[i])) i++;
        if (i == start) return false;
        col = ColToIndex(s.Substring(start, i - start));

        if (i < s.Length && s[i] == '$') i++;
        start = i;
        while (i < s.Length && char.IsDigit(s[i])) i++;
        if (i == start) return false;

        row = int.Parse(s.Substring(start, i - start), CultureInfo.InvariantCulture);
        return i == s.Length;
    }

    public static bool TryParseColumn(string s, out int col)
    {
        col = 0;
        if (string.IsNullOrEmpty(s)) return false;

        int i = 0;
        if (s[i] == '$') i++;

        int start = i;
        while (i < s.Length && IsAsciiLetter(s[i])) i++;
        if (i == start || i != s.Length) return false;

        col = ColToIndex(s.Substring(start, i - start));
        return true;
    }

    public static bool TryParseRow(string s, out int row)
    {
        row = 0;
        if (string.IsNullOrEmpty(s)) return false;

        int i = 0;
        if (s[i] == '$') i++;

        int start = i;
        while (i < s.Length && char.IsDigit(s[i])) i++;
        if (i == start || i != s.Length) return false;

        row = int.Parse(s.Substring(start, i - start), CultureInfo.InvariantCulture);
        return true;
    }

    public static bool TryParseRange(string s, out int r1, out int c1, out int r2, out int c2)
    {
        r1 = 0; c1 = 0; r2 = 0; c2 = 0;
        if (string.IsNullOrEmpty(s)) return false;

        int bang = s.LastIndexOf('!');
        if (bang >= 0) s = s.Substring(bang + 1);

        int colon = s.IndexOf(':');
        if (colon < 0)
        {
            if (!TryParseRef(s, out r1, out c1)) return false;
            r2 = r1;
            c2 = c1;
            return true;
        }

        string left = s.Substring(0, colon);
        string right = s.Substring(colon + 1);

        if (TryParseRef(left, out r1, out c1) && TryParseRef(right, out r2, out c2))
        {
            if (r1 > r2) { int t = r1; r1 = r2; r2 = t; }
            if (c1 > c2) { int t = c1; c1 = c2; c2 = t; }
            return true;
        }

        if (TryParseColumn(left, out c1) && TryParseColumn(right, out c2))
        {
            r1 = 0;
            r2 = 0;
            if (c1 > c2) { int t = c1; c1 = c2; c2 = t; }
            return true;
        }

        if (TryParseRow(left, out r1) && TryParseRow(right, out r2))
        {
            c1 = 0;
            c2 = 0;
            if (r1 > r2) { int t = r1; r1 = r2; r2 = t; }
            return true;
        }

        return false;
    }
}
