using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

interface IBook
{
    List<SheetInfo> Sheets { get; }
    List<DefinedName> DefinedNames { get; }
    bool Date1904 { get; }
    bool HasVba { get; }
    string SheetName(int id);
    string GetFormatCode(int styleIndex);
    string GetDimension(SheetInfo sh);
    void ReadCells(SheetInfo sh, int r1, int c1, int r2, int c2, CellSink sink);
    byte[] GetVbaProject();
    DefinedName FindName(string name, string sheetName);
    bool TryResolveNameRange(string name, string sheetName,
        out string targetSheet, out int r1, out int c1, out int r2, out int c2);
    void Dispose();
}

static class Book
{
    public static IBook Open(string path)
    {
        byte[] head = new byte[8];
        using (FileStream fs = File.OpenRead(path))
        {
            fs.ReadExactly(head, 0, 8);
        }

        // ZIP: xlsx / xlsm / xlsb
        if (head[0] == 0x50 && head[1] == 0x4B && head[2] == 0x03 && head[3] == 0x04)
        {
            using (ZipArchive z = ZipFile.OpenRead(path))
            {
                for (int i = 0; i < z.Entries.Count; i++)
                {
                    if (string.Equals(z.Entries[i].FullName, "xl/workbook.bin", StringComparison.OrdinalIgnoreCase))
                        return new XlsbBook(path);
                }
            }
            return new Xlsx(path);
        }

        // CFB: xls (BIFF8)
        if (head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0)
            return new XlsBook(path);

        throw new Exception("unsupported file format (expected xlsx / xlsm / xlsb / xls)");
    }
}

// Shared name-resolution logic for the binary readers. Xlsx has its own
// (equivalent) copy, kept separate to avoid disturbing the working OOXML path.
abstract class BookBase : IBook
{
    protected List<DefinedName> definedNames = new List<DefinedName>();
    protected List<string> allSheetNames = new List<string>();

    public List<DefinedName> DefinedNames { get { return definedNames; } }

    public abstract List<SheetInfo> Sheets { get; }
    public abstract bool Date1904 { get; }
    public abstract bool HasVba { get; }
    public abstract string GetFormatCode(int styleIndex);
    public abstract string GetDimension(SheetInfo sh);
    public abstract void ReadCells(SheetInfo sh, int r1, int c1, int r2, int c2, CellSink sink);
    public abstract byte[] GetVbaProject();
    public abstract void Dispose();

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
            targetSheet = null;
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
}
