using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

// ---------------------------------------------------------------- values

enum EvalKind
{
    Number,
    Text,
    Bool,
    Error,
    Empty,
    Range
}

class EvalValue
{
    public EvalKind Kind;
    public double Number;
    public string Text;
    public bool Bool;
    public string Error;
    public EvalValue[,] Cells;
    public int Rows;
    public int Cols;
    public int StartRow;
    public int StartCol;

    public static readonly EvalValue Empty = MakeEmpty();
    public static readonly EvalValue True = MakeBool(true);
    public static readonly EvalValue False = MakeBool(false);

    static EvalValue MakeEmpty()
    {
        EvalValue v = new EvalValue();
        v.Kind = EvalKind.Empty;
        return v;
    }

    static EvalValue MakeBool(bool b)
    {
        EvalValue v = new EvalValue();
        v.Kind = EvalKind.Bool;
        v.Bool = b;
        return v;
    }

    public static EvalValue Num(double d)
    {
        EvalValue v = new EvalValue();
        v.Kind = EvalKind.Number;
        v.Number = d;
        return v;
    }

    public static EvalValue Str(string s)
    {
        EvalValue v = new EvalValue();
        v.Kind = EvalKind.Text;
        v.Text = s == null ? "" : s;
        return v;
    }

    public static EvalValue BoolValue(bool b)
    {
        return b ? True : False;
    }

    public static EvalValue Err(string code)
    {
        EvalValue v = new EvalValue();
        v.Kind = EvalKind.Error;
        v.Error = code == null ? "#VALUE!" : code;
        return v;
    }

    public static EvalValue RangeValue(EvalValue[,] cells)
    {
        EvalValue v = new EvalValue();
        v.Kind = EvalKind.Range;
        v.Cells = cells;
        v.Rows = cells.GetLength(0);
        v.Cols = cells.GetLength(1);
        return v;
    }

    public bool IsError
    {
        get { return Kind == EvalKind.Error; }
    }

    public EvalValue Scalar()
    {
        if (Kind != EvalKind.Range) return this;
        if (Rows <= 0 || Cols <= 0) return Empty;
        return Cells[0, 0];
    }

    public IEnumerable<EvalValue> Flatten()
    {
        if (Kind != EvalKind.Range)
        {
            yield return this;
            yield break;
        }
        for (int r = 0; r < Rows; r++)
            for (int c = 0; c < Cols; c++)
                yield return Cells[r, c];
    }

    public double ToNumber(out bool ok)
    {
        ok = true;
        switch (Kind)
        {
            case EvalKind.Number:
                return Number;
            case EvalKind.Bool:
                return Bool ? 1.0 : 0.0;
            case EvalKind.Empty:
                return 0.0;
            case EvalKind.Text:
                {
                    string s = Text.Trim();
                    if (s.Length == 0) return 0.0;
                    double d;
                    if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                        return d;
                    ok = false;
                    return 0.0;
                }
            case EvalKind.Range:
                return Scalar().ToNumber(out ok);
            default:
                ok = false;
                return 0.0;
        }
    }

    public string ToText()
    {
        switch (Kind)
        {
            case EvalKind.Number:
                return Number.ToString("G15", CultureInfo.InvariantCulture);
            case EvalKind.Text:
                return Text;
            case EvalKind.Bool:
                return Bool ? "TRUE" : "FALSE";
            case EvalKind.Empty:
                return "";
            case EvalKind.Error:
                return Error;
            case EvalKind.Range:
                return Scalar().ToText();
            default:
                return "";
        }
    }

    public bool ToBool(out bool ok)
    {
        ok = true;
        switch (Kind)
        {
            case EvalKind.Bool:
                return Bool;
            case EvalKind.Number:
                return Number != 0.0;
            case EvalKind.Empty:
                return false;
            case EvalKind.Text:
                {
                    string s = Text.Trim();
                    if (string.Equals(s, "TRUE", StringComparison.OrdinalIgnoreCase)) return true;
                    if (string.Equals(s, "FALSE", StringComparison.OrdinalIgnoreCase)) return false;
                    double d;
                    if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                        return d != 0.0;
                    ok = false;
                    return false;
                }
            case EvalKind.Range:
                return Scalar().ToBool(out ok);
            default:
                ok = false;
                return false;
        }
    }

    public string ToDisplayString()
    {
        switch (Kind)
        {
            case EvalKind.Number:
                return Number.ToString("G15", CultureInfo.InvariantCulture);
            case EvalKind.Text:
                return Text;
            case EvalKind.Bool:
                return Bool ? "TRUE" : "FALSE";
            case EvalKind.Empty:
                return "";
            case EvalKind.Error:
                return Error;
            case EvalKind.Range:
                return Scalar().ToDisplayString();
            default:
                return "";
        }
    }

    public int CompareTo(EvalValue other)
    {
        EvalValue a = Scalar();
        EvalValue b = other.Scalar();

        if (a.Kind == EvalKind.Text && b.Kind == EvalKind.Text)
            return string.Compare(a.Text, b.Text, StringComparison.OrdinalIgnoreCase);

        bool oka, okb;
        double da = a.ToNumber(out oka);
        double db = b.ToNumber(out okb);
        if (oka && okb)
            return da < db ? -1 : (da > db ? 1 : 0);

        return string.Compare(a.ToText(), b.ToText(), StringComparison.OrdinalIgnoreCase);
    }

    public string TypeCode()
    {
        switch (Kind)
        {
            case EvalKind.Number: return "n";
            case EvalKind.Text: return "s";
            case EvalKind.Bool: return "b";
            case EvalKind.Error: return "e";
            case EvalKind.Empty: return "z";
            case EvalKind.Range: return "a";
            default: return "z";
        }
    }
}

// ---------------------------------------------------------------- cells

class EvalCell
{
    public string V;
    public string T;
    public string F;
    public int Style;
    public string Sheet;
    public int Row;
    public int Col;
}

class EvalRangeRef
{
    public string Sheet;
    public int R1;
    public int C1;
    public int R2;
    public int C2;
}

class LoadedRange
{
    public string Sheet;
    public int R1;
    public int C1;
    public int R2;
    public int C2;
}

class EvalDepNode
{
    public string Key;
    public string Sheet;
    public int Row;
    public int Col;
    public EvalNode Ast;
    public List<string> Dependencies = new List<string>();
    public List<string> Dependents = new List<string>();
}

class EvalContext
{
    IBook book;
    string defaultSheet;
    Dictionary<string, EvalCell> cells = new Dictionary<string, EvalCell>(StringComparer.OrdinalIgnoreCase);
    Dictionary<string, EvalValue> values = new Dictionary<string, EvalValue>(StringComparer.OrdinalIgnoreCase);
    HashSet<string> evaluating = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    List<LoadedRange> loadedRanges = new List<LoadedRange>();
    Dictionary<string, EvalDepNode> nodes = new Dictionary<string, EvalDepNode>(StringComparer.OrdinalIgnoreCase);
    Dictionary<string, EvalNode> astCache = new Dictionary<string, EvalNode>(StringComparer.OrdinalIgnoreCase);

    public EvalContext(IBook book, string defaultSheet)
    {
        this.book = book;
        if (string.IsNullOrEmpty(defaultSheet) && book.Sheets.Count > 0)
            defaultSheet = book.Sheets[0].Name;
        this.defaultSheet = defaultSheet;
    }

    public bool Date1904 { get { return book.Date1904; } }
    public string DefaultSheet { get { return defaultSheet; } }

    public EvalValue Evaluate(string expression)
    {
        if (expression == null) throw new EvalParseException("empty expression");
        string f = expression.Trim();
        if (f.StartsWith("=")) f = f.Substring(1);
        if (f.Length == 0) throw new EvalParseException("empty expression");

        EvalNode root = EvalParser.Parse(f);
        Prepare(root, defaultSheet);
        return root.Eval(this, defaultSheet, 0, 0);
    }

    public EvalValue EvaluateFormula(string formula, string sheet, int row, int col)
    {
        if (formula == null) throw new EvalParseException("empty expression");
        string f = formula.Trim();
        if (f.StartsWith("=")) f = f.Substring(1);
        if (f.Length == 0) throw new EvalParseException("empty expression");
        EvalNode node = EvalParser.Parse(f);
        return node.Eval(this, sheet, row, col);
    }

    public EvalValue GetCell(string sheet, int row, int col)
    {
        sheet = NormalizeSheet(sheet);
        if (sheet == null) return EvalValue.Err("#REF!");
        SheetInfo sh = FindSheet(sheet);
        if (sh == null) return EvalValue.Err("#REF!");
        return GetCellNormalized(sh.Name, row, col);
    }

    EvalValue GetCellNormalized(string sheet, int row, int col)
    {
        string key = CellKey(sheet, row, col);
        EvalValue cached;
        if (values.TryGetValue(key, out cached)) return cached;

        EvalCell cell;
        if (!cells.TryGetValue(key, out cell))
        {
            if (!IsCovered(sheet, row, col, row, col))
                EnsureRange(sheet, row, col, row, col);
            cells.TryGetValue(key, out cell);
        }
        if (cell == null) return EvalValue.Empty;

        if (cell.F != null && cell.F.Length > 0 && (cell.V == null || cell.V.Length == 0))
        {
            if (evaluating.Contains(key)) return EvalValue.Err("#REF!");
            evaluating.Add(key);
            try
            {
                EvalValue v = EvaluateFormula(cell.F, sheet, row, col);
                v = v.Scalar();
                values[key] = v;
                return v;
            }
            catch (EvalParseException)
            {
                values[key] = EvalValue.Err("#NAME?");
                return values[key];
            }
            finally
            {
                evaluating.Remove(key);
            }
        }

        EvalValue r = CellToValue(cell);
        values[key] = r;
        return r;
    }

    public EvalValue GetRange(string sheet, int r1, int c1, int r2, int c2)
    {
        sheet = NormalizeSheet(sheet);
        if (sheet == null) return EvalValue.Err("#REF!");
        SheetInfo sh = FindSheet(sheet);
        if (sh == null) return EvalValue.Err("#REF!");
        sheet = sh.Name;

        if (r1 > r2) { int t = r1; r1 = r2; r2 = t; }
        if (c1 > c2) { int t = c1; c1 = c2; c2 = t; }

        if (!IsCovered(sheet, r1, c1, r2, c2))
            EnsureRange(sheet, r1, c1, r2, c2);
        int rows = r2 - r1 + 1;
        int cols = c2 - c1 + 1;
        EvalValue[,] arr = new EvalValue[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                arr[r, c] = GetCellNormalized(sheet, r1 + r, c1 + c);
        EvalValue range = EvalValue.RangeValue(arr);
        range.StartRow = r1;
        range.StartCol = c1;
        return range;
    }

    public EvalValue ResolveName(string name, string sheet)
    {
        string target;
        int r1, c1, r2, c2;
        if (book.TryResolveNameRange(name, sheet, out target, out r1, out c1, out r2, out c2))
        {
            if (target == null) target = sheet;
            return GetRange(target, r1, c1, r2, c2);
        }

        DefinedName dn = book.FindName(name, sheet);
        if (dn != null && dn.Ref != null)
        {
            string refText = dn.Ref.Trim();
            if (refText.StartsWith("=")) refText = refText.Substring(1);
            if (refText.Length > 0 && refText.IndexOf('!') < 0)
            {
                int a, b, c, d;
                if (!A1.TryParseRange(refText, out a, out b, out c, out d))
                {
                    try
                    {
                        return EvaluateFormula(refText, sheet, 0, 0).Scalar();
                    }
                    catch (EvalParseException)
                    {
                    }
                }
            }
        }
        return EvalValue.Err("#NAME?");
    }

    public bool TryResolveNameBounds(string name, string sheet,
        out string targetSheet, out int r1, out int c1, out int r2, out int c2)
    {
        return book.TryResolveNameRange(name, sheet, out targetSheet, out r1, out c1, out r2, out c2);
    }

    void EnsureRange(string sheet, int r1, int c1, int r2, int c2)
    {
        if (IsCovered(sheet, r1, c1, r2, c2)) return;

        SheetInfo sh = FindSheet(sheet);
        if (sh == null) return;

        book.ReadCells(sh, r1, c1, r2, c2, delegate (string sn, int row, int col, string val, string type, string f, int styleIndex)
        {
            EvalCell c = new EvalCell();
            c.V = val;
            c.T = type;
            c.F = f;
            c.Style = styleIndex;
            c.Sheet = sheet;
            c.Row = row;
            c.Col = col;
            cells[CellKey(sheet, row, col)] = c;
        });

        LoadedRange range = new LoadedRange();
        range.Sheet = sheet;
        range.R1 = r1;
        range.C1 = c1;
        range.R2 = r2;
        range.C2 = c2;
        loadedRanges.Add(range);
    }

    bool IsCovered(string sheet, int r1, int c1, int r2, int c2)
    {
        for (int i = 0; i < loadedRanges.Count; i++)
        {
            LoadedRange r = loadedRanges[i];
            if (!string.Equals(r.Sheet, sheet, StringComparison.OrdinalIgnoreCase)) continue;
            if (r1 >= r.R1 && c1 >= r.C1 && r2 <= r.R2 && c2 <= r.C2) return true;
        }
        return false;
    }

    void Prepare(EvalNode root, string rootSheet)
    {
        nodes.Clear();
        astCache.Clear();

        List<EvalRangeRef> refs = new List<EvalRangeRef>();
        root.CollectRefs(this, rootSheet, refs);
        LoadRanges(refs);
        BuildFormulaGraph(refs);
        EvaluateGraph();
    }

    void LoadRanges(List<EvalRangeRef> refs)
    {
        Dictionary<string, List<EvalRangeRef>> grouped = new Dictionary<string, List<EvalRangeRef>>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < refs.Count; i++)
        {
            EvalRangeRef r = refs[i];
            string sheet = NormalizeSheet(r.Sheet);
            if (sheet == null) continue;

            List<EvalRangeRef> list;
            if (!grouped.TryGetValue(sheet, out list))
            {
                list = new List<EvalRangeRef>();
                grouped[sheet] = list;
            }

            EvalRangeRef copy = new EvalRangeRef();
            copy.Sheet = sheet;
            copy.R1 = r.R1;
            copy.C1 = r.C1;
            copy.R2 = r.R2;
            copy.C2 = r.C2;
            list.Add(copy);
        }

        foreach (KeyValuePair<string, List<EvalRangeRef>> pair in grouped)
        {
            List<EvalRangeRef> merged = MergeRanges(pair.Value);
            for (int i = 0; i < merged.Count; i++)
            {
                EvalRangeRef r = merged[i];
                EnsureRange(r.Sheet, r.R1, r.C1, r.R2, r.C2);
            }
        }
    }

    static List<EvalRangeRef> MergeRanges(List<EvalRangeRef> ranges)
    {
        ranges.Sort(delegate (EvalRangeRef a, EvalRangeRef b)
        {
            int cmp = a.R1.CompareTo(b.R1);
            if (cmp != 0) return cmp;
            return a.C1.CompareTo(b.C1);
        });

        List<EvalRangeRef> result = new List<EvalRangeRef>();
        EvalRangeRef current = null;

        for (int i = 0; i < ranges.Count; i++)
        {
            EvalRangeRef r = ranges[i];
            if (current == null)
            {
                current = CloneRange(r);
                continue;
            }

            bool adjacentRows = r.R1 <= current.R2 + 1 && r.R2 >= current.R1 - 1;
            bool adjacentCols = r.C1 <= current.C2 + 1 && r.C2 >= current.C1 - 1;
            if (adjacentRows && adjacentCols)
            {
                if (r.R1 < current.R1) current.R1 = r.R1;
                if (r.C1 < current.C1) current.C1 = r.C1;
                if (r.R2 > current.R2) current.R2 = r.R2;
                if (r.C2 > current.C2) current.C2 = r.C2;
            }
            else
            {
                result.Add(current);
                current = CloneRange(r);
            }
        }

        if (current != null) result.Add(current);
        return result;
    }

    static EvalRangeRef CloneRange(EvalRangeRef r)
    {
        EvalRangeRef copy = new EvalRangeRef();
        copy.Sheet = r.Sheet;
        copy.R1 = r.R1;
        copy.C1 = r.C1;
        copy.R2 = r.R2;
        copy.C2 = r.C2;
        return copy;
    }

    void AddFormulaCells(List<EvalRangeRef> refs, Queue<EvalCell> pending, HashSet<string> seen)
    {
        for (int i = 0; i < refs.Count; i++)
        {
            EvalRangeRef r = refs[i];
            string sheet = NormalizeSheet(r.Sheet);
            if (sheet == null) continue;

            for (int row = r.R1; row <= r.R2; row++)
            {
                for (int col = r.C1; col <= r.C2; col++)
                {
                    string key = CellKey(sheet, row, col);
                    EvalCell cell;
                    if (!cells.TryGetValue(key, out cell)) continue;
                    if (cell.F == null || cell.F.Length == 0) continue;
                    if (cell.V != null && cell.V.Length > 0) continue;
                    if (!seen.Add(key)) continue;

                    EvalDepNode node = new EvalDepNode();
                    node.Key = key;
                    node.Sheet = sheet;
                    node.Row = row;
                    node.Col = col;
                    nodes[key] = node;
                    pending.Enqueue(cell);
                }
            }
        }
    }

    void BuildFormulaGraph(List<EvalRangeRef> rootRefs)
    {
        Queue<EvalCell> pending = new Queue<EvalCell>();
        HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        AddFormulaCells(rootRefs, pending, seen);

        while (pending.Count > 0)
        {
            EvalCell cell = pending.Dequeue();
            string key = CellKey(cell.Sheet, cell.Row, cell.Col);
            EvalNode ast;
            if (!astCache.TryGetValue(key, out ast))
            {
                try
                {
                    ast = EvalParser.Parse(cell.F);
                }
                catch (EvalParseException)
                {
                    ast = null;
                }
                astCache[key] = ast;
            }

            EvalDepNode node;
            if (nodes.TryGetValue(key, out node)) node.Ast = ast;

            List<EvalRangeRef> refs = new List<EvalRangeRef>();
            if (ast != null) ast.CollectRefs(this, cell.Sheet, refs);
            LoadRanges(refs);
            AddFormulaCells(refs, pending, seen);
        }

        foreach (EvalDepNode node in nodes.Values)
        {
            EvalNode ast;
            if (!astCache.TryGetValue(node.Key, out ast) || ast == null) continue;

            List<EvalRangeRef> refs = new List<EvalRangeRef>();
            ast.CollectRefs(this, node.Sheet, refs);
            HashSet<string> deps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < refs.Count; i++)
            {
                EvalRangeRef r = refs[i];
                string sheet = NormalizeSheet(r.Sheet);
                if (sheet == null) continue;

                for (int row = r.R1; row <= r.R2; row++)
                {
                    for (int col = r.C1; col <= r.C2; col++)
                    {
                        string depKey = CellKey(sheet, row, col);
                        if (nodes.ContainsKey(depKey)) deps.Add(depKey);
                    }
                }
            }

            node.Dependencies.Clear();
            node.Dependencies.AddRange(deps);
        }

        foreach (EvalDepNode node in nodes.Values)
        {
            for (int i = 0; i < node.Dependencies.Count; i++)
            {
                EvalDepNode dep;
                if (nodes.TryGetValue(node.Dependencies[i], out dep))
                    dep.Dependents.Add(node.Key);
            }
        }
    }

    void EvaluateGraph()
    {
        if (nodes.Count == 0) return;

        Dictionary<string, int> indegree = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        Queue<string> ready = new Queue<string>();

        foreach (EvalDepNode node in nodes.Values)
        {
            indegree[node.Key] = node.Dependencies.Count;
            if (node.Dependencies.Count == 0) ready.Enqueue(node.Key);
        }

        HashSet<string> evaluated = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        while (ready.Count > 0)
        {
            string key = ready.Dequeue();
            EvalDepNode node = nodes[key];
            EvalValue value;

            if (node.Ast == null)
            {
                value = EvalValue.Err("#NAME?");
            }
            else
            {
                try
                {
                    value = node.Ast.Eval(this, node.Sheet, node.Row, node.Col).Scalar();
                }
                catch (EvalParseException)
                {
                    value = EvalValue.Err("#NAME?");
                }
            }

            values[key] = value;
            evaluated.Add(key);

            for (int i = 0; i < node.Dependents.Count; i++)
            {
                string dependent = node.Dependents[i];
                indegree[dependent]--;
                if (indegree[dependent] == 0) ready.Enqueue(dependent);
            }
        }

        foreach (EvalDepNode node in nodes.Values)
        {
            if (!evaluated.Contains(node.Key))
                values[node.Key] = EvalValue.Err("#REF!");
        }
    }

    EvalValue CellToValue(EvalCell cell)
    {
        if (cell.V == null) return EvalValue.Empty;
        if (cell.T == "n")
        {
            double d;
            if (double.TryParse(cell.V, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                return EvalValue.Num(d);
            return EvalValue.Str(cell.V);
        }
        if (cell.T == "s") return EvalValue.Str(cell.V);
        if (cell.T == "b") return EvalValue.BoolValue(cell.V == "true");
        if (cell.T == "e") return EvalValue.Err(cell.V);
        if (cell.T == "d") return EvalValue.Str(cell.V);

        double n;
        if (double.TryParse(cell.V, NumberStyles.Float, CultureInfo.InvariantCulture, out n))
            return EvalValue.Num(n);
        return EvalValue.Str(cell.V);
    }

    string NormalizeSheet(string sheet)
    {
        if (string.IsNullOrEmpty(sheet)) sheet = defaultSheet;
        if (string.IsNullOrEmpty(sheet)) return null;
        SheetInfo sh = FindSheet(sheet);
        return sh != null ? sh.Name : sheet;
    }

    SheetInfo FindSheet(string name)
    {
        for (int i = 0; i < book.Sheets.Count; i++)
        {
            SheetInfo sh = book.Sheets[i];
            if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase)) return sh;
        }
        return null;
    }

    static string CellKey(string sheet, int row, int col)
    {
        return sheet + "!" + row.ToString(CultureInfo.InvariantCulture) + "," + col.ToString(CultureInfo.InvariantCulture);
    }

    static string RangeKey(string sheet, int r1, int c1, int r2, int c2)
    {
        return sheet + "!" + r1.ToString(CultureInfo.InvariantCulture) + "," + c1.ToString(CultureInfo.InvariantCulture)
            + ":" + r2.ToString(CultureInfo.InvariantCulture) + "," + c2.ToString(CultureInfo.InvariantCulture);
    }
}

// ---------------------------------------------------------------- parser

class EvalParseException : Exception
{
    public EvalParseException(string message) : base(message) { }
}

enum EvalTokenType
{
    Number,
    String,
    Ident,
    Error,
    Op,
    LParen,
    RParen,
    Comma,
    Colon,
    Bang,
    End
}

class EvalToken
{
    public EvalTokenType Type;
    public string Text;
    public double Number;

    public EvalToken(EvalTokenType type, string text)
    {
        Type = type;
        Text = text;
    }
}

static class EvalParser
{
    public static EvalNode Parse(string expression)
    {
        List<EvalToken> tokens = Lex(expression);
        EvalParserImpl p = new EvalParserImpl(tokens);
        EvalNode n = p.ParseComparison();
        if (p.Current.Type != EvalTokenType.End)
            throw new EvalParseException("unexpected token: " + p.Current.Text);
        return n;
    }

    static List<EvalToken> Lex(string s)
    {
        List<EvalToken> tokens = new List<EvalToken>();
        int i = 0;

        while (i < s.Length)
        {
            char c = s[i];

            if (char.IsWhiteSpace(c)) { i++; continue; }

            if (c == '"')
            {
                i++;
                StringBuilder sb = new StringBuilder();
                while (i < s.Length)
                {
                    if (s[i] == '"')
                    {
                        if (i + 1 < s.Length && s[i + 1] == '"')
                        {
                            sb.Append('"');
                            i += 2;
                            continue;
                        }
                        i++;
                        break;
                    }
                    sb.Append(s[i]);
                    i++;
                }
                tokens.Add(new EvalToken(EvalTokenType.String, sb.ToString()));
                continue;
            }

            if (c == '\'')
            {
                i++;
                StringBuilder sb = new StringBuilder();
                while (i < s.Length)
                {
                    if (s[i] == '\'')
                    {
                        if (i + 1 < s.Length && s[i + 1] == '\'')
                        {
                            sb.Append('\'');
                            i += 2;
                            continue;
                        }
                        i++;
                        break;
                    }
                    sb.Append(s[i]);
                    i++;
                }
                tokens.Add(new EvalToken(EvalTokenType.Ident, sb.ToString()));
                continue;
            }

            if (c == '#')
            {
                int start = i;
                i++;
                while (i < s.Length)
                {
                    char d = s[i];
                    if (char.IsLetterOrDigit(d) || d == '/' || d == '!' || d == '?' || d == '_' || d == '.')
                    {
                        i++;
                        continue;
                    }
                    break;
                }
                tokens.Add(new EvalToken(EvalTokenType.Error, s.Substring(start, i - start)));
                continue;
            }

            if (char.IsDigit(c) || (c == '.' && i + 1 < s.Length && char.IsDigit(s[i + 1])))
            {
                int start = i;
                while (i < s.Length && char.IsDigit(s[i])) i++;
                if (i < s.Length && s[i] == '.')
                {
                    i++;
                    while (i < s.Length && char.IsDigit(s[i])) i++;
                }
                if (i < s.Length && (s[i] == 'e' || s[i] == 'E'))
                {
                    i++;
                    if (i < s.Length && (s[i] == '+' || s[i] == '-')) i++;
                    while (i < s.Length && char.IsDigit(s[i])) i++;
                }
                string numText = s.Substring(start, i - start);
                EvalToken t = new EvalToken(EvalTokenType.Number, numText);
                t.Number = double.Parse(numText, NumberStyles.Float, CultureInfo.InvariantCulture);
                tokens.Add(t);
                continue;
            }

            if (char.IsLetter(c) || c == '_' || c == '$' || c == '.')
            {
                int start = i;
                while (i < s.Length)
                {
                    char d = s[i];
                    if (char.IsLetterOrDigit(d) || d == '_' || d == '$' || d == '.') i++;
                    else break;
                }
                tokens.Add(new EvalToken(EvalTokenType.Ident, s.Substring(start, i - start)));
                continue;
            }

            if (c == '<' && i + 1 < s.Length && s[i + 1] == '=')
            {
                tokens.Add(new EvalToken(EvalTokenType.Op, "<="));
                i += 2;
                continue;
            }
            if (c == '>' && i + 1 < s.Length && s[i + 1] == '=')
            {
                tokens.Add(new EvalToken(EvalTokenType.Op, ">="));
                i += 2;
                continue;
            }
            if (c == '<' && i + 1 < s.Length && s[i + 1] == '>')
            {
                tokens.Add(new EvalToken(EvalTokenType.Op, "<>"));
                i += 2;
                continue;
            }
            if (c == '=' && i + 1 < s.Length && s[i + 1] == '=')
            {
                tokens.Add(new EvalToken(EvalTokenType.Op, "="));
                i += 2;
                continue;
            }

            if (c == '+' || c == '-' || c == '*' || c == '/' || c == '^' || c == '&'
                || c == '=' || c == '<' || c == '>' || c == '%')
            {
                tokens.Add(new EvalToken(EvalTokenType.Op, c.ToString()));
                i++;
                continue;
            }

            if (c == '(') { tokens.Add(new EvalToken(EvalTokenType.LParen, "(")); i++; continue; }
            if (c == ')') { tokens.Add(new EvalToken(EvalTokenType.RParen, ")")); i++; continue; }
            if (c == ',') { tokens.Add(new EvalToken(EvalTokenType.Comma, ",")); i++; continue; }
            if (c == ':') { tokens.Add(new EvalToken(EvalTokenType.Colon, ":")); i++; continue; }
            if (c == '!') { tokens.Add(new EvalToken(EvalTokenType.Bang, "!")); i++; continue; }

            throw new EvalParseException("unexpected character: " + c);
        }

        tokens.Add(new EvalToken(EvalTokenType.End, ""));
        return tokens;
    }
}

class EvalParserImpl
{
    List<EvalToken> tokens;
    int pos;

    public EvalParserImpl(List<EvalToken> tokens)
    {
        this.tokens = tokens;
        this.pos = 0;
    }

    public EvalToken Current { get { return tokens[pos]; } }

    void Advance() { if (pos < tokens.Count - 1) pos++; }

    bool MatchOp(string op)
    {
        if (Current.Type == EvalTokenType.Op && Current.Text == op)
        {
            Advance();
            return true;
        }
        return false;
    }

    EvalToken Expect(EvalTokenType type)
    {
        if (Current.Type != type)
            throw new EvalParseException("expected " + type + " but got " + Current.Text);
        EvalToken t = Current;
        Advance();
        return t;
    }

    static bool IsComparison(string op)
    {
        return op == "=" || op == "<>" || op == "<" || op == "<=" || op == ">" || op == ">=";
    }

    public EvalNode ParseComparison()
    {
        EvalNode left = ParseConcat();
        while (Current.Type == EvalTokenType.Op && IsComparison(Current.Text))
        {
            string op = Current.Text;
            Advance();
            EvalNode right = ParseConcat();
            left = new EvalBinaryNode(op, left, right);
        }
        return left;
    }

    EvalNode ParseConcat()
    {
        EvalNode left = ParseAdd();
        while (MatchOp("&"))
        {
            EvalNode right = ParseAdd();
            left = new EvalBinaryNode("&", left, right);
        }
        return left;
    }

    EvalNode ParseAdd()
    {
        EvalNode left = ParseMul();
        while (Current.Type == EvalTokenType.Op && (Current.Text == "+" || Current.Text == "-"))
        {
            string op = Current.Text;
            Advance();
            EvalNode right = ParseMul();
            left = new EvalBinaryNode(op, left, right);
        }
        return left;
    }

    EvalNode ParseMul()
    {
        EvalNode left = ParsePow();
        while (Current.Type == EvalTokenType.Op && (Current.Text == "*" || Current.Text == "/"))
        {
            string op = Current.Text;
            Advance();
            EvalNode right = ParsePow();
            left = new EvalBinaryNode(op, left, right);
        }
        return left;
    }

    EvalNode ParsePow()
    {
        EvalNode left = ParseUnary();
        while (MatchOp("^"))
        {
            EvalNode right = ParseUnary();
            left = new EvalBinaryNode("^", left, right);
        }
        return left;
    }

    EvalNode ParseUnary()
    {
        if (Current.Type == EvalTokenType.Op && (Current.Text == "+" || Current.Text == "-"))
        {
            string op = Current.Text;
            Advance();
            return new EvalUnaryNode(op, ParseUnary());
        }
        return ParsePostfix();
    }

    EvalNode ParsePostfix()
    {
        EvalNode node = ParsePrimary();
        while (MatchOp("%"))
            node = new EvalUnaryNode("%", node);
        return node;
    }

    EvalNode ParsePrimary()
    {
        if (Current.Type == EvalTokenType.Number)
        {
            double d = Current.Number;
            Advance();
            return new EvalLiteralNode(EvalValue.Num(d));
        }

        if (Current.Type == EvalTokenType.String)
        {
            string s = Current.Text;
            Advance();
            return new EvalLiteralNode(EvalValue.Str(s));
        }

        if (Current.Type == EvalTokenType.Error)
        {
            string e = Current.Text;
            Advance();
            return new EvalLiteralNode(EvalValue.Err(e));
        }

        if (Current.Type == EvalTokenType.Ident)
        {
            string name = Current.Text;
            Advance();

            if (Current.Type == EvalTokenType.LParen)
                return ParseFunction(name);

            if (Current.Type == EvalTokenType.Bang)
            {
                Advance();
                return ParseRefOrName(name);
            }

            if (string.Equals(name, "TRUE", StringComparison.OrdinalIgnoreCase))
                return new EvalLiteralNode(EvalValue.True);
            if (string.Equals(name, "FALSE", StringComparison.OrdinalIgnoreCase))
                return new EvalLiteralNode(EvalValue.False);

            return ParseRefOrName(null, name);
        }

        if (Current.Type == EvalTokenType.LParen)
        {
            Advance();
            EvalNode n = ParseComparison();
            Expect(EvalTokenType.RParen);
            return n;
        }

        throw new EvalParseException("unexpected token: " + Current.Text);
    }

    EvalNode ParseFunction(string name)
    {
        Expect(EvalTokenType.LParen);
        List<EvalNode> args = new List<EvalNode>();

        if (Current.Type != EvalTokenType.RParen)
        {
            while (true)
            {
                if (Current.Type == EvalTokenType.Comma)
                {
                    args.Add(new EvalLiteralNode(EvalValue.Empty));
                    Advance();
                    if (Current.Type == EvalTokenType.RParen)
                    {
                        args.Add(new EvalLiteralNode(EvalValue.Empty));
                        break;
                    }
                    continue;
                }

                args.Add(ParseComparison());

                if (Current.Type == EvalTokenType.Comma)
                {
                    Advance();
                    if (Current.Type == EvalTokenType.RParen)
                    {
                        args.Add(new EvalLiteralNode(EvalValue.Empty));
                        break;
                    }
                    continue;
                }
                break;
            }
        }

        Expect(EvalTokenType.RParen);
        return new EvalFunctionNode(name, args);
    }

    EvalNode ParseRefOrName(string sheet)
    {
        return ParseRefOrName(sheet, null);
    }

    EvalNode ParseRefOrName(string sheet, string knownName)
    {
        string name;
        if (knownName != null) name = knownName;
        else
        {
            EvalToken t = Expect(EvalTokenType.Ident);
            name = t.Text;
        }

        int r1, c1, r2, c2;
        if (A1.TryParseRef(name, out r1, out c1))
        {
            if (Current.Type == EvalTokenType.Colon)
            {
                Advance();
                EvalToken t2 = Expect(EvalTokenType.Ident);
                if (!A1.TryParseRef(t2.Text, out r2, out c2))
                    throw new EvalParseException("bad range end: " + t2.Text);
                return new EvalRefNode(sheet, r1, c1, r2, c2, null);
            }
            return new EvalRefNode(sheet, r1, c1, r1, c1, null);
        }

        return new EvalRefNode(sheet, 0, 0, 0, 0, name);
    }
}

// ---------------------------------------------------------------- nodes

abstract class EvalNode
{
    public abstract EvalValue Eval(EvalContext ctx, string sheet, int row, int col);
    public abstract void CollectRefs(EvalContext ctx, string currentSheet, List<EvalRangeRef> refs);
}

class EvalLiteralNode : EvalNode
{
    EvalValue value;
    public EvalLiteralNode(EvalValue value) { this.value = value; }
    public override EvalValue Eval(EvalContext ctx, string sheet, int row, int col) { return value; }
    public override void CollectRefs(EvalContext ctx, string currentSheet, List<EvalRangeRef> refs) { }
}

class EvalRefNode : EvalNode
{
    string sheet;
    int r1, c1, r2, c2;
    string name;

    public EvalRefNode(string sheet, int r1, int c1, int r2, int c2, string name)
    {
        this.sheet = sheet;
        this.r1 = r1;
        this.c1 = c1;
        this.r2 = r2;
        this.c2 = c2;
        this.name = name;
    }

    public override EvalValue Eval(EvalContext ctx, string currentSheet, int row, int col)
    {
        string targetSheet = sheet != null ? sheet : currentSheet;
        if (name != null) return ctx.ResolveName(name, targetSheet);
        if (r1 == r2 && c1 == c2) return ctx.GetCell(targetSheet, r1, c1);
        return ctx.GetRange(targetSheet, r1, c1, r2, c2);
    }

    public override void CollectRefs(EvalContext ctx, string currentSheet, List<EvalRangeRef> refs)
    {
        string targetSheet;
        int br1, bc1, br2, bc2;
        if (!TryGetBounds(ctx, currentSheet, out targetSheet, out br1, out bc1, out br2, out bc2)) return;

        EvalRangeRef r = new EvalRangeRef();
        r.Sheet = targetSheet;
        r.R1 = br1;
        r.C1 = bc1;
        r.R2 = br2;
        r.C2 = bc2;
        refs.Add(r);
    }

    public bool TryGetBounds(EvalContext ctx, string currentSheet,
        out string targetSheet, out int br1, out int bc1, out int br2, out int bc2)
    {
        if (name != null)
            return ctx.TryResolveNameBounds(name, sheet != null ? sheet : currentSheet,
                out targetSheet, out br1, out bc1, out br2, out bc2);

        targetSheet = sheet != null ? sheet : currentSheet;
        br1 = r1;
        bc1 = c1;
        br2 = r2;
        bc2 = c2;
        return true;
    }
}

class EvalUnaryNode : EvalNode
{
    string op;
    EvalNode arg;

    public EvalUnaryNode(string op, EvalNode arg)
    {
        this.op = op;
        this.arg = arg;
    }

    public override void CollectRefs(EvalContext ctx, string currentSheet, List<EvalRangeRef> refs)
    {
        arg.CollectRefs(ctx, currentSheet, refs);
    }

    public override EvalValue Eval(EvalContext ctx, string sheet, int row, int col)
    {
        EvalValue v = arg.Eval(ctx, sheet, row, col);
        if (v.IsError) return v;
        v = v.Scalar();

        if (op == "+" || op == "-")
        {
            bool ok;
            double d = v.ToNumber(out ok);
            if (!ok) return EvalValue.Err("#VALUE!");
            return EvalValue.Num(op == "-" ? -d : d);
        }
        if (op == "%")
        {
            bool ok;
            double d = v.ToNumber(out ok);
            if (!ok) return EvalValue.Err("#VALUE!");
            return EvalValue.Num(d / 100.0);
        }
        return EvalValue.Err("#VALUE!");
    }
}

class EvalBinaryNode : EvalNode
{
    string op;
    EvalNode left;
    EvalNode right;

    public EvalBinaryNode(string op, EvalNode left, EvalNode right)
    {
        this.op = op;
        this.left = left;
        this.right = right;
    }

    public override void CollectRefs(EvalContext ctx, string currentSheet, List<EvalRangeRef> refs)
    {
        left.CollectRefs(ctx, currentSheet, refs);
        right.CollectRefs(ctx, currentSheet, refs);
    }

    public override EvalValue Eval(EvalContext ctx, string sheet, int row, int col)
    {
        EvalValue a = left.Eval(ctx, sheet, row, col);
        if (a.IsError) return a;
        EvalValue b = right.Eval(ctx, sheet, row, col);
        if (b.IsError) return b;

        if (op == "&")
            return EvalValue.Str(a.ToText() + b.ToText());

        if (op == "=" || op == "<>" || op == "<" || op == "<=" || op == ">" || op == ">=")
        {
            int cmp = Compare(a, b);
            bool r;
            if (op == "=") r = cmp == 0;
            else if (op == "<>") r = cmp != 0;
            else if (op == "<") r = cmp < 0;
            else if (op == "<=") r = cmp <= 0;
            else if (op == ">") r = cmp > 0;
            else r = cmp >= 0;
            return EvalValue.BoolValue(r);
        }

        bool oka, okb;
        double da = a.ToNumber(out oka);
        double db = b.ToNumber(out okb);
        if (!oka || !okb) return EvalValue.Err("#VALUE!");

        if (op == "+") return EvalValue.Num(da + db);
        if (op == "-") return EvalValue.Num(da - db);
        if (op == "*") return EvalValue.Num(da * db);
        if (op == "/")
        {
            if (db == 0.0) return EvalValue.Err("#DIV/0!");
            return EvalValue.Num(da / db);
        }
        if (op == "^") return EvalValue.Num(Math.Pow(da, db));

        return EvalValue.Err("#VALUE!");
    }

    static int Compare(EvalValue a, EvalValue b)
    {
        return a.CompareTo(b);
    }
}

static class EvalWildcard
{
    public static bool HasWildcard(string s)
    {
        return s.IndexOf('*') >= 0 || s.IndexOf('?') >= 0;
    }

    public static bool Match(string pattern, string text)
    {
        int p = 0;
        int t = 0;
        int star = -1;
        int match = 0;

        while (t < text.Length)
        {
            if (p < pattern.Length && (pattern[p] == '?' || char.ToUpperInvariant(pattern[p]) == char.ToUpperInvariant(text[t])))
            {
                p++;
                t++;
            }
            else if (p < pattern.Length && pattern[p] == '*')
            {
                star = p;
                p++;
                match = t;
            }
            else if (star >= 0)
            {
                p = star + 1;
                match++;
                t = match;
            }
            else
            {
                return false;
            }
        }

        while (p < pattern.Length && pattern[p] == '*') p++;
        return p == pattern.Length;
    }
}

class EvalCriteria
{
    public string Op;
    public EvalValue Value;
    public bool Wildcard;

    public static EvalCriteria From(EvalValue value)
    {
        EvalCriteria c = new EvalCriteria();
        c.Op = "=";
        c.Value = value.Scalar();

        if (c.Value.Kind == EvalKind.Text)
        {
            string s = c.Value.Text;
            string op = null;
            if (s.StartsWith(">=", StringComparison.Ordinal)) { op = ">="; s = s.Substring(2); }
            else if (s.StartsWith("<=", StringComparison.Ordinal)) { op = "<="; s = s.Substring(2); }
            else if (s.StartsWith("<>", StringComparison.Ordinal)) { op = "<>"; s = s.Substring(2); }
            else if (s.StartsWith(">", StringComparison.Ordinal)) { op = ">"; s = s.Substring(1); }
            else if (s.StartsWith("<", StringComparison.Ordinal)) { op = "<"; s = s.Substring(1); }
            else if (s.StartsWith("=", StringComparison.Ordinal)) { op = "="; s = s.Substring(1); }

            if (op != null) c.Op = op;
            c.Value = ParseOperand(s);
        }

        c.Wildcard = c.Value.Kind == EvalKind.Text && EvalWildcard.HasWildcard(c.Value.Text);
        return c;
    }

    static EvalValue ParseOperand(string s)
    {
        if (s.Length == 0) return EvalValue.Empty;
        double d;
        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
            return EvalValue.Num(d);
        return EvalValue.Str(s);
    }

    public bool Match(EvalValue cell)
    {
        cell = cell.Scalar();
        if (cell.IsError) return false;

        bool cellEmpty = cell.Kind == EvalKind.Empty || (cell.Kind == EvalKind.Text && cell.Text.Length == 0);

        if (Op == "=" && Value.Kind == EvalKind.Empty) return cellEmpty;
        if (Op == "<>" && Value.Kind == EvalKind.Empty) return !cellEmpty;

        if (Wildcard && Value.Kind == EvalKind.Text && cell.Kind == EvalKind.Text)
        {
            bool m = EvalWildcard.Match(Value.Text, cell.Text);
            return Op == "<>" ? !m : m;
        }

        int cmp = cell.CompareTo(Value);
        if (Op == "=") return cmp == 0;
        if (Op == "<>") return cmp != 0;
        if (Op == "<") return cmp < 0;
        if (Op == "<=") return cmp <= 0;
        if (Op == ">") return cmp > 0;
        if (Op == ">=") return cmp >= 0;
        return false;
    }
}

class EvalFunctionNode : EvalNode
{
    string name;
    List<EvalNode> args;

    public EvalFunctionNode(string name, List<EvalNode> args)
    {
        this.name = name;
        this.args = args;
    }

    public override void CollectRefs(EvalContext ctx, string currentSheet, List<EvalRangeRef> refs)
    {
        for (int i = 0; i < args.Count; i++)
            args[i].CollectRefs(ctx, currentSheet, refs);
    }

    EvalValue Arg(EvalContext ctx, string sheet, int row, int col, int i)
    {
        return args[i].Eval(ctx, sheet, row, col);
    }

    List<EvalValue> FlattenArgs(EvalContext ctx, string sheet, int row, int col)
    {
        List<EvalValue> vals = new List<EvalValue>();
        for (int i = 0; i < args.Count; i++)
        {
            EvalValue v = Arg(ctx, sheet, row, col, i);
            if (v.Kind == EvalKind.Range)
            {
                foreach (EvalValue x in v.Flatten()) vals.Add(x);
            }
            else
            {
                vals.Add(v);
            }
        }
        return vals;
    }

    static bool HasError(List<EvalValue> vals)
    {
        for (int i = 0; i < vals.Count; i++)
            if (vals[i].IsError) return true;
        return false;
    }

    public override EvalValue Eval(EvalContext ctx, string sheet, int row, int col)
    {
        string fn = name.ToUpperInvariant();
        if (fn.StartsWith("_XLFN.")) fn = fn.Substring(6);

        if (fn == "IF")
        {
            if (args.Count < 2) return EvalValue.Err("#VALUE!");
            EvalValue cond = Arg(ctx, sheet, row, col, 0);
            if (cond.IsError) return cond;
            bool ok;
            bool b = cond.ToBool(out ok);
            if (!ok) return EvalValue.Err("#VALUE!");
            if (b) return Arg(ctx, sheet, row, col, 1);
            if (args.Count >= 3) return Arg(ctx, sheet, row, col, 2);
            return EvalValue.False;
        }

        if (fn == "IFERROR")
        {
            if (args.Count < 2) return EvalValue.Err("#VALUE!");
            EvalValue v = Arg(ctx, sheet, row, col, 0);
            if (v.IsError) return Arg(ctx, sheet, row, col, 1);
            return v;
        }

        if (fn == "TRUE") return EvalValue.True;
        if (fn == "FALSE") return EvalValue.False;
        if (fn == "NA") return EvalValue.Err("#N/A");
        if (fn == "PI") return EvalValue.Num(Math.PI);

        if (fn == "ADDRESS")
        {
            if (args.Count < 2) return EvalValue.Err("#VALUE!");

            EvalValue rowValue = Arg(ctx, sheet, row, col, 0);
            if (rowValue.IsError) return rowValue;
            EvalValue colValue = Arg(ctx, sheet, row, col, 1);
            if (colValue.IsError) return colValue;

            bool rowOk, colOk;
            int rowNum = (int)rowValue.ToNumber(out rowOk);
            int colNum = (int)colValue.ToNumber(out colOk);
            if (!rowOk || !colOk || rowNum < 1 || colNum < 1) return EvalValue.Err("#VALUE!");

            int absNum = 1;
            if (args.Count >= 3)
            {
                EvalValue absValue = Arg(ctx, sheet, row, col, 2);
                if (absValue.IsError) return absValue;
                bool absOk;
                absNum = (int)absValue.ToNumber(out absOk);
                if (!absOk || absNum < 1 || absNum > 4) return EvalValue.Err("#VALUE!");
            }

            bool useA1 = true;
            if (args.Count >= 4)
            {
                EvalValue a1Value = Arg(ctx, sheet, row, col, 3);
                if (a1Value.IsError) return a1Value;
                bool a1Ok;
                useA1 = a1Value.ToBool(out a1Ok);
                if (!a1Ok) return EvalValue.Err("#VALUE!");
            }

            string sheetText = null;
            if (args.Count >= 5)
            {
                EvalValue sheetValue = Arg(ctx, sheet, row, col, 4);
                if (sheetValue.IsError) return sheetValue;
                sheetText = sheetValue.ToText();
            }

            string cell;
            if (useA1)
            {
                string colName = A1.IndexToCol(colNum);
                string rowText = rowNum.ToString(CultureInfo.InvariantCulture);
                if (absNum == 1) cell = "$" + colName + "$" + rowText;
                else if (absNum == 2) cell = colName + "$" + rowText;
                else if (absNum == 3) cell = "$" + colName + rowText;
                else cell = colName + rowText;
            }
            else
            {
                string rowText = rowNum.ToString(CultureInfo.InvariantCulture);
                string colText = colNum.ToString(CultureInfo.InvariantCulture);
                if (absNum == 1) cell = "R" + rowText + "C" + colText;
                else if (absNum == 2) cell = "R" + rowText + "C[" + colText + "]";
                else if (absNum == 3) cell = "R[" + rowText + "]C" + colText;
                else cell = "R[" + rowText + "]C[" + colText + "]";
            }

            if (!string.IsNullOrEmpty(sheetText))
                cell = QuoteSheet(sheetText) + "!" + cell;
            return EvalValue.Str(cell);
        }

        if (fn == "ROW" || fn == "COLUMN")
        {
            if (args.Count == 0)
            {
                int index = fn == "ROW" ? row : col;
                if (index <= 0) return EvalValue.Err("#VALUE!");
                return EvalValue.Num(index);
            }
            if (args.Count > 1) return EvalValue.Err("#VALUE!");

            if (args[0] is EvalRefNode)
            {
                EvalRefNode refNode = (EvalRefNode)args[0];
                string targetSheet;
                int r1, c1, r2, c2;
                if (!refNode.TryGetBounds(ctx, sheet, out targetSheet, out r1, out c1, out r2, out c2))
                    return EvalValue.Err("#NAME?");
                return EvalValue.Num(fn == "ROW" ? r1 : c1);
            }

            EvalValue v = Arg(ctx, sheet, row, col, 0);
            if (v.IsError) return v;
            if (v.Kind == EvalKind.Range)
                return EvalValue.Num(fn == "ROW" ? v.StartRow : v.StartCol);
            return EvalValue.Err("#VALUE!");
        }

        if (fn == "ROWS" || fn == "COLUMNS")
        {
            if (args.Count < 1) return EvalValue.Err("#VALUE!");
            EvalValue v = Arg(ctx, sheet, row, col, 0);
            if (v.IsError) return v;
            if (v.Kind == EvalKind.Range)
                return EvalValue.Num(fn == "ROWS" ? v.Rows : v.Cols);
            return EvalValue.Num(1);
        }

        if (fn == "SUBSTITUTE")
        {
            if (args.Count < 3) return EvalValue.Err("#VALUE!");

            EvalValue textValue = Arg(ctx, sheet, row, col, 0);
            if (textValue.IsError) return textValue;
            EvalValue oldValue = Arg(ctx, sheet, row, col, 1);
            if (oldValue.IsError) return oldValue;
            EvalValue newValue = Arg(ctx, sheet, row, col, 2);
            if (newValue.IsError) return newValue;

            string text = textValue.ToText();
            string oldText = oldValue.ToText();
            string newText = newValue.ToText();

            if (oldText.Length == 0) return EvalValue.Str(text);

            int instance = 0;
            if (args.Count >= 4)
            {
                EvalValue instanceValue = Arg(ctx, sheet, row, col, 3);
                if (instanceValue.IsError) return instanceValue;
                bool instanceOk;
                instance = (int)instanceValue.ToNumber(out instanceOk);
                if (!instanceOk || instance < 1) return EvalValue.Err("#VALUE!");
            }

            if (instance == 0) return EvalValue.Str(text.Replace(oldText, newText));

            int from = 0;
            int count = 0;
            while (from <= text.Length)
            {
                int index = text.IndexOf(oldText, from, StringComparison.Ordinal);
                if (index < 0) break;
                count++;
                if (count == instance)
                    return EvalValue.Str(text.Substring(0, index) + newText + text.Substring(index + oldText.Length));
                from = index + oldText.Length;
            }
            return EvalValue.Str(text);
        }

        if (fn == "REPLACE")
        {
            if (args.Count < 4) return EvalValue.Err("#VALUE!");

            EvalValue textValue = Arg(ctx, sheet, row, col, 0);
            if (textValue.IsError) return textValue;
            EvalValue startValue = Arg(ctx, sheet, row, col, 1);
            if (startValue.IsError) return startValue;
            EvalValue countValue = Arg(ctx, sheet, row, col, 2);
            if (countValue.IsError) return countValue;
            EvalValue newValue = Arg(ctx, sheet, row, col, 3);
            if (newValue.IsError) return newValue;

            bool startOk, countOk;
            int start = (int)startValue.ToNumber(out startOk);
            int count = (int)countValue.ToNumber(out countOk);
            if (!startOk || !countOk || start < 1 || count < 0) return EvalValue.Err("#VALUE!");

            string text = textValue.ToText();
            string newText = newValue.ToText();
            int start0 = start - 1;
            if (start0 > text.Length) start0 = text.Length;
            if (count > text.Length - start0) count = text.Length - start0;
            return EvalValue.Str(text.Substring(0, start0) + newText + text.Substring(start0 + count));
        }

        if (fn == "INDIRECT")
        {
            if (args.Count < 1) return EvalValue.Err("#VALUE!");
            EvalValue textValue = Arg(ctx, sheet, row, col, 0);
            if (textValue.IsError) return textValue;

            if (args.Count >= 2)
            {
                EvalValue a1 = Arg(ctx, sheet, row, col, 1);
                if (a1.IsError) return a1;
                bool a1ok;
                bool useA1 = a1.ToBool(out a1ok);
                if (!a1ok || !useA1) return EvalValue.Err("#REF!");
            }

            string refText = textValue.ToText().Trim();
            if (refText.Length == 0 || refText.StartsWith("[", StringComparison.Ordinal))
                return EvalValue.Err("#REF!");

            try
            {
                EvalNode node = EvalParser.Parse(refText);
                if (node is EvalRefNode)
                    return node.Eval(ctx, sheet, row, col);
            }
            catch (EvalParseException)
            {
            }
            return EvalValue.Err("#REF!");
        }

        if (fn == "MATCH")
        {
            if (args.Count < 2) return EvalValue.Err("#VALUE!");
            EvalValue lookup = Arg(ctx, sheet, row, col, 0);
            if (lookup.IsError) return lookup;
            EvalValue arrayValue = Arg(ctx, sheet, row, col, 1);
            if (arrayValue.IsError) return arrayValue;
            if (arrayValue.Kind != EvalKind.Range) return EvalValue.Err("#N/A");
            if (arrayValue.Rows > 1 && arrayValue.Cols > 1) return EvalValue.Err("#N/A");

            int matchType = 1;
            if (args.Count >= 3)
            {
                EvalValue mt = Arg(ctx, sheet, row, col, 2);
                if (mt.IsError) return mt;
                bool mtok;
                double d = mt.ToNumber(out mtok);
                if (!mtok) return EvalValue.Err("#VALUE!");
                matchType = (int)d;
            }

            List<EvalValue> flat = new List<EvalValue>();
            for (int r = 0; r < arrayValue.Rows; r++)
                for (int c = 0; c < arrayValue.Cols; c++)
                    flat.Add(arrayValue.Cells[r, c]);

            if (matchType == 0)
            {
                for (int i = 0; i < flat.Count; i++)
                    if (MatchLookup(flat[i], lookup)) return EvalValue.Num(i + 1);
                return EvalValue.Err("#N/A");
            }

            int found = -1;
            for (int i = 0; i < flat.Count; i++)
            {
                int cmp = flat[i].CompareTo(lookup);
                if (matchType > 0)
                {
                    if (cmp <= 0) found = i;
                }
                else
                {
                    if (cmp >= 0) found = i;
                }
            }
            if (found < 0) return EvalValue.Err("#N/A");
            return EvalValue.Num(found + 1);
        }

        if (fn == "INDEX")
        {
            if (args.Count < 2) return EvalValue.Err("#VALUE!");
            EvalValue arrayValue = Arg(ctx, sheet, row, col, 0);
            if (arrayValue.IsError) return arrayValue;
            if (arrayValue.Kind != EvalKind.Range) return EvalValue.Err("#VALUE!");

            EvalValue rowValue = Arg(ctx, sheet, row, col, 1);
            if (rowValue.IsError) return rowValue;
            bool rowOk;
            double rowD = rowValue.ToNumber(out rowOk);
            if (!rowOk) return EvalValue.Err("#VALUE!");
            int rowIndex = (int)rowD;

            int colIndex = 1;
            bool hasCol = false;
            if (args.Count >= 3)
            {
                EvalValue colValue = Arg(ctx, sheet, row, col, 2);
                if (colValue.IsError) return colValue;
                bool colOk;
                double colD = colValue.ToNumber(out colOk);
                if (!colOk) return EvalValue.Err("#VALUE!");
                colIndex = (int)colD;
                hasCol = true;
            }

            if (arrayValue.Rows == 1 && arrayValue.Cols > 1 && !hasCol)
            {
                colIndex = rowIndex;
                rowIndex = 1;
            }
            else if (arrayValue.Cols == 1)
            {
                colIndex = 1;
            }
            else if (!hasCol)
            {
                return EvalValue.Err("#VALUE!");
            }

            if (rowIndex == 0 && colIndex == 0) return arrayValue;

            if (rowIndex == 0)
            {
                if (colIndex < 1 || colIndex > arrayValue.Cols) return EvalValue.Err("#REF!");
                EvalValue[,] colArr = new EvalValue[arrayValue.Rows, 1];
                for (int r = 0; r < arrayValue.Rows; r++)
                    colArr[r, 0] = arrayValue.Cells[r, colIndex - 1];
                return EvalValue.RangeValue(colArr);
            }

            if (colIndex == 0)
            {
                if (rowIndex < 1 || rowIndex > arrayValue.Rows) return EvalValue.Err("#REF!");
                EvalValue[,] rowArr = new EvalValue[1, arrayValue.Cols];
                for (int c = 0; c < arrayValue.Cols; c++)
                    rowArr[0, c] = arrayValue.Cells[rowIndex - 1, c];
                return EvalValue.RangeValue(rowArr);
            }

            if (rowIndex < 1 || rowIndex > arrayValue.Rows || colIndex < 1 || colIndex > arrayValue.Cols)
                return EvalValue.Err("#REF!");

            return arrayValue.Cells[rowIndex - 1, colIndex - 1];
        }

        if (fn == "VLOOKUP")
        {
            if (args.Count < 3) return EvalValue.Err("#VALUE!");
            EvalValue lookup = Arg(ctx, sheet, row, col, 0);
            if (lookup.IsError) return lookup;
            EvalValue table = Arg(ctx, sheet, row, col, 1);
            if (table.IsError) return table;
            if (table.Kind != EvalKind.Range) return EvalValue.Err("#VALUE!");

            EvalValue colValue = Arg(ctx, sheet, row, col, 2);
            if (colValue.IsError) return colValue;
            bool colOk;
            double colD = colValue.ToNumber(out colOk);
            if (!colOk) return EvalValue.Err("#VALUE!");
            int colIndex = (int)colD;
            if (colIndex < 1) return EvalValue.Err("#VALUE!");
            if (colIndex > table.Cols) return EvalValue.Err("#REF!");

            bool rangeLookup = true;
            if (args.Count >= 4)
            {
                EvalValue rv = Arg(ctx, sheet, row, col, 3);
                if (rv.IsError) return rv;
                bool rvOk;
                bool b = rv.ToBool(out rvOk);
                if (!rvOk) return EvalValue.Err("#VALUE!");
                rangeLookup = b;
            }

            int found = -1;
            if (rangeLookup)
            {
                for (int r = 0; r < table.Rows; r++)
                {
                    if (table.Cells[r, 0].CompareTo(lookup) <= 0) found = r;
                }
            }
            else
            {
                for (int r = 0; r < table.Rows; r++)
                {
                    if (MatchLookup(table.Cells[r, 0], lookup))
                    {
                        found = r;
                        break;
                    }
                }
            }

            if (found < 0) return EvalValue.Err("#N/A");
            return table.Cells[found, colIndex - 1];
        }

        if (fn == "SUMIF" || fn == "COUNTIF")
        {
            if (args.Count < 2) return EvalValue.Err("#VALUE!");
            EvalValue rangeValue = Arg(ctx, sheet, row, col, 0);
            if (rangeValue.IsError) return rangeValue;
            if (rangeValue.Kind != EvalKind.Range) return EvalValue.Err("#VALUE!");

            EvalValue criteriaValue = Arg(ctx, sheet, row, col, 1);
            if (criteriaValue.IsError) return criteriaValue;
            EvalCriteria criteria = EvalCriteria.From(criteriaValue);

            EvalValue sumRange = rangeValue;
            if (fn == "SUMIF" && args.Count >= 3)
            {
                sumRange = Arg(ctx, sheet, row, col, 2);
                if (sumRange.IsError) return sumRange;
                if (sumRange.Kind != EvalKind.Range) return EvalValue.Err("#VALUE!");
                if (sumRange.Rows != rangeValue.Rows || sumRange.Cols != rangeValue.Cols)
                    return EvalValue.Err("#VALUE!");
            }

            double sum = 0.0;
            int count = 0;
            for (int r = 0; r < rangeValue.Rows; r++)
            {
                for (int c = 0; c < rangeValue.Cols; c++)
                {
                    if (!criteria.Match(rangeValue.Cells[r, c])) continue;
                    count++;
                    if (fn == "SUMIF")
                    {
                        EvalValue sv = sumRange.Cells[r, c].Scalar();
                        if (sv.IsError) continue;
                        if (sv.Kind == EvalKind.Number) sum += sv.Number;
                    }
                }
            }

            if (fn == "SUMIF") return EvalValue.Num(sum);
            return EvalValue.Num(count);
        }

        List<EvalValue> vals = FlattenArgs(ctx, sheet, row, col);

        if (fn == "SUM" || fn == "AVERAGE" || fn == "MIN" || fn == "MAX" || fn == "COUNT")
        {
            if (HasError(vals)) return FirstError(vals);
            double sum = 0.0;
            int count = 0;
            double min = 0.0;
            double max = 0.0;
            bool has = false;

            for (int i = 0; i < vals.Count; i++)
            {
                EvalValue v = vals[i].Scalar();
                if (v.Kind == EvalKind.Number)
                {
                    sum += v.Number;
                    count++;
                    if (!has) { min = v.Number; max = v.Number; has = true; }
                    else { if (v.Number < min) min = v.Number; if (v.Number > max) max = v.Number; }
                }
            }

            if (fn == "SUM") return EvalValue.Num(sum);
            if (fn == "AVERAGE")
            {
                if (count == 0) return EvalValue.Err("#DIV/0!");
                return EvalValue.Num(sum / count);
            }
            if (fn == "COUNT") return EvalValue.Num(count);
            if (fn == "MIN") return EvalValue.Num(has ? min : 0.0);
            if (fn == "MAX") return EvalValue.Num(has ? max : 0.0);
        }

        if (fn == "COUNTA")
        {
            int count = 0;
            for (int i = 0; i < vals.Count; i++)
                if (vals[i].Kind != EvalKind.Empty) count++;
            return EvalValue.Num(count);
        }

        if (fn == "AND" || fn == "OR")
        {
            if (HasError(vals)) return FirstError(vals);
            bool result = fn == "AND";
            for (int i = 0; i < vals.Count; i++)
            {
                EvalValue v = vals[i].Scalar();
                if (v.Kind == EvalKind.Empty) continue;
                bool ok;
                bool b = v.ToBool(out ok);
                if (!ok) return EvalValue.Err("#VALUE!");
                if (fn == "AND") result = result && b;
                else result = result || b;
            }
            return EvalValue.BoolValue(result);
        }

        if (fn == "NOT")
        {
            if (vals.Count < 1) return EvalValue.Err("#VALUE!");
            if (vals[0].IsError) return vals[0];
            bool ok;
            bool b = vals[0].ToBool(out ok);
            if (!ok) return EvalValue.Err("#VALUE!");
            return EvalValue.BoolValue(!b);
        }

        if (fn == "ABS" || fn == "INT" || fn == "SQRT")
        {
            if (vals.Count < 1) return EvalValue.Err("#VALUE!");
            if (vals[0].IsError) return vals[0];
            bool ok;
            double d = vals[0].ToNumber(out ok);
            if (!ok) return EvalValue.Err("#VALUE!");
            if (fn == "ABS") return EvalValue.Num(Math.Abs(d));
            if (fn == "INT") return EvalValue.Num(Math.Floor(d));
            if (d < 0) return EvalValue.Err("#NUM!");
            return EvalValue.Num(Math.Sqrt(d));
        }

        if (fn == "MOD" || fn == "POWER")
        {
            if (vals.Count < 2) return EvalValue.Err("#VALUE!");
            if (vals[0].IsError) return vals[0];
            if (vals[1].IsError) return vals[1];
            bool ok1, ok2;
            double a = vals[0].ToNumber(out ok1);
            double b = vals[1].ToNumber(out ok2);
            if (!ok1 || !ok2) return EvalValue.Err("#VALUE!");
            if (fn == "MOD")
            {
                if (b == 0.0) return EvalValue.Err("#DIV/0!");
                return EvalValue.Num(a - b * Math.Floor(a / b));
            }
            return EvalValue.Num(Math.Pow(a, b));
        }

        if (fn == "ROUND" || fn == "ROUNDUP" || fn == "ROUNDDOWN")
        {
            if (vals.Count < 2) return EvalValue.Err("#VALUE!");
            if (vals[0].IsError) return vals[0];
            if (vals[1].IsError) return vals[1];
            bool ok1, ok2;
            double v = vals[0].ToNumber(out ok1);
            double d = vals[1].ToNumber(out ok2);
            if (!ok1 || !ok2) return EvalValue.Err("#VALUE!");
            int digits = (int)d;
            return EvalValue.Num(RoundValue(v, digits, fn));
        }

        if (fn == "LEN" || fn == "TRIM" || fn == "UPPER" || fn == "LOWER" || fn == "VALUE")
        {
            if (vals.Count < 1) return EvalValue.Err("#VALUE!");
            if (vals[0].IsError) return vals[0];
            string s = vals[0].ToText();
            if (fn == "LEN") return EvalValue.Num(s.Length);
            if (fn == "TRIM") return EvalValue.Str(TrimExcel(s));
            if (fn == "UPPER") return EvalValue.Str(s.ToUpperInvariant());
            if (fn == "LOWER") return EvalValue.Str(s.ToLowerInvariant());
            double d;
            if (double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                return EvalValue.Num(d);
            return EvalValue.Err("#VALUE!");
        }

        if (fn == "LEFT" || fn == "RIGHT" || fn == "MID")
        {
            if (vals.Count < 1) return EvalValue.Err("#VALUE!");
            if (vals[0].IsError) return vals[0];
            string s = vals[0].ToText();

            if (fn == "LEFT" || fn == "RIGHT")
            {
                int n = 1;
                if (vals.Count >= 2)
                {
                    if (vals[1].IsError) return vals[1];
                    bool ok;
                    double d = vals[1].ToNumber(out ok);
                    if (!ok) return EvalValue.Err("#VALUE!");
                    n = (int)d;
                }
                if (n < 0) return EvalValue.Err("#VALUE!");
                if (n > s.Length) n = s.Length;
                if (fn == "LEFT") return EvalValue.Str(s.Substring(0, n));
                return EvalValue.Str(s.Substring(s.Length - n, n));
            }

            if (vals.Count < 3) return EvalValue.Err("#VALUE!");
            if (vals[1].IsError) return vals[1];
            if (vals[2].IsError) return vals[2];
            bool ok1, ok2;
            double startD = vals[1].ToNumber(out ok1);
            double lenD = vals[2].ToNumber(out ok2);
            if (!ok1 || !ok2) return EvalValue.Err("#VALUE!");
            int start = (int)startD;
            int length = (int)lenD;
            if (start < 1 || length < 0) return EvalValue.Err("#VALUE!");
            int from = start - 1;
            if (from >= s.Length) return EvalValue.Str("");
            if (length > s.Length - from) length = s.Length - from;
            return EvalValue.Str(s.Substring(from, length));
        }

        if (fn == "CONCATENATE")
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < vals.Count; i++)
            {
                if (vals[i].IsError) return vals[i];
                sb.Append(vals[i].ToText());
            }
            return EvalValue.Str(sb.ToString());
        }

        if (fn == "TEXT")
        {
            if (vals.Count < 2) return EvalValue.Err("#VALUE!");
            if (vals[0].IsError) return vals[0];
            if (vals[1].IsError) return vals[1];
            string code = vals[1].ToText();
            EvalValue v = vals[0].Scalar();
            if (v.Kind == EvalKind.Number)
                return EvalValue.Str(NumFmt.Format(v.Number, code, ctx.Date1904));
            if (v.Kind == EvalKind.Text)
                return EvalValue.Str(NumFmt.FormatText(v.Text, code));
            return EvalValue.Str(v.ToText());
        }

        if (fn == "ISBLANK" || fn == "ISNUMBER" || fn == "ISTEXT" || fn == "ISERROR" || fn == "ISNA")
        {
            if (vals.Count < 1) return EvalValue.Err("#VALUE!");
            EvalValue v = vals[0].Scalar();
            if (fn == "ISBLANK") return EvalValue.BoolValue(v.Kind == EvalKind.Empty);
            if (fn == "ISNUMBER") return EvalValue.BoolValue(v.Kind == EvalKind.Number);
            if (fn == "ISTEXT") return EvalValue.BoolValue(v.Kind == EvalKind.Text);
            if (fn == "ISERROR") return EvalValue.BoolValue(v.IsError);
            return EvalValue.BoolValue(v.IsError && v.Error == "#N/A");
        }

        if (fn == "DATE")
        {
            if (vals.Count < 3) return EvalValue.Err("#VALUE!");
            if (vals[0].IsError) return vals[0];
            if (vals[1].IsError) return vals[1];
            if (vals[2].IsError) return vals[2];
            bool ok1, ok2, ok3;
            int year = (int)vals[0].ToNumber(out ok1);
            int month = (int)vals[1].ToNumber(out ok2);
            int day = (int)vals[2].ToNumber(out ok3);
            if (!ok1 || !ok2 || !ok3) return EvalValue.Err("#VALUE!");
            try
            {
                DateTime dt = new DateTime(year, month, day);
                return EvalValue.Num(ToSerial(dt, ctx.Date1904));
            }
            catch
            {
                return EvalValue.Err("#NUM!");
            }
        }

        if (fn == "YEAR" || fn == "MONTH" || fn == "DAY")
        {
            if (vals.Count < 1) return EvalValue.Err("#VALUE!");
            if (vals[0].IsError) return vals[0];
            bool ok;
            double serial = vals[0].ToNumber(out ok);
            if (!ok) return EvalValue.Err("#VALUE!");
            DateTime dt = FromSerial(serial, ctx.Date1904);
            if (fn == "YEAR") return EvalValue.Num(dt.Year);
            if (fn == "MONTH") return EvalValue.Num(dt.Month);
            return EvalValue.Num(dt.Day);
        }

        return EvalValue.Err("#NAME?");
    }

    static string QuoteSheet(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        bool quote = false;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (!char.IsLetterOrDigit(c) && c != '_')
            {
                quote = true;
                break;
            }
        }
        if (!quote) return s;
        return "'" + s.Replace("'", "''") + "'";
    }

    static bool MatchLookup(EvalValue cell, EvalValue lookup)
    {
        cell = cell.Scalar();
        lookup = lookup.Scalar();
        if (cell.IsError) return false;
        if (lookup.Kind == EvalKind.Text && EvalWildcard.HasWildcard(lookup.Text) && cell.Kind == EvalKind.Text)
            return EvalWildcard.Match(lookup.Text, cell.Text);
        return cell.CompareTo(lookup) == 0;
    }

    static EvalValue FirstError(List<EvalValue> vals)
    {
        for (int i = 0; i < vals.Count; i++)
            if (vals[i].IsError) return vals[i];
        return EvalValue.Err("#VALUE!");
    }

    static double RoundValue(double v, int digits, string mode)
    {
        double factor = Math.Pow(10.0, digits);
        double scaled = v * factor;
        double rounded;
        if (mode == "ROUND")
            rounded = Math.Round(scaled, 0, MidpointRounding.AwayFromZero);
        else if (mode == "ROUNDUP")
            rounded = scaled >= 0 ? Math.Ceiling(scaled) : Math.Floor(scaled);
        else
            rounded = scaled >= 0 ? Math.Floor(scaled) : Math.Ceiling(scaled);
        return rounded / factor;
    }

    static string TrimExcel(string s)
    {
        StringBuilder sb = new StringBuilder();
        bool lastSpace = false;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == ' ')
            {
                if (sb.Length > 0 && !lastSpace) sb.Append(' ');
                lastSpace = true;
            }
            else
            {
                sb.Append(c);
                lastSpace = false;
            }
        }
        if (sb.Length > 0 && sb[sb.Length - 1] == ' ') sb.Length--;
        return sb.ToString();
    }

    static double ToSerial(DateTime dt, bool date1904)
    {
        DateTime epoch = date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
        double d = (dt - epoch).TotalDays;
        if (!date1904 && dt < new DateTime(1900, 3, 1)) d -= 1.0;
        return d;
    }

    static DateTime FromSerial(double serial, bool date1904)
    {
        DateTime epoch = date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
        double d = serial;
        if (!date1904 && serial < 61.0) d += 1.0;
        return epoch.AddDays(d);
    }
}

// ---------------------------------------------------------------- JSON

static class EvalJson
{
    public static void AppendValue(StringBuilder sb, EvalValue v)
    {
        switch (v.Kind)
        {
            case EvalKind.Number:
                if (double.IsNaN(v.Number) || double.IsInfinity(v.Number)) sb.Append("null");
                else sb.Append(v.Number.ToString("R", CultureInfo.InvariantCulture));
                break;
            case EvalKind.Text:
                AppendString(sb, v.Text);
                break;
            case EvalKind.Bool:
                sb.Append(v.Bool ? "true" : "false");
                break;
            case EvalKind.Error:
                AppendString(sb, v.Error);
                break;
            case EvalKind.Empty:
                sb.Append("null");
                break;
            case EvalKind.Range:
                sb.Append('[');
                for (int r = 0; r < v.Rows; r++)
                {
                    if (r > 0) sb.Append(',');
                    sb.Append('[');
                    for (int c = 0; c < v.Cols; c++)
                    {
                        if (c > 0) sb.Append(',');
                        AppendValue(sb, v.Cells[r, c]);
                    }
                    sb.Append(']');
                }
                sb.Append(']');
                break;
        }
    }

    public static void AppendString(StringBuilder sb, string s)
    {
        sb.Append('"');
        if (s != null)
        {
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
                else sb.Append(c);
            }
        }
        sb.Append('"');
    }
}
