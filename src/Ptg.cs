using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

// BIFF8 (and later BIFF12) formula token decompiler.
// Token layout follows MS-XLS 2.5.198; function table from xlrd (BSD).
delegate byte[] SharedFormulaLookupEx(int row, int col, out int baseRow, out int baseCol);

class FormulaContext
{
    public string[] SheetNames;
    public List<DefinedName> Names;
    public Func<int, string> ExternSheet;
    public bool RelDelta;
    public int BaseRow;
    public int BaseCol;
    public int RowOffset;
    public int ColOffset;
    public Func<int, int, byte[]> SharedFormula;
    public SharedFormulaLookupEx SharedFormulaEx;
    public int CellRow;
    public int CellCol;
}

class PtgOperand
{
    public int Rank;
    public string Text;

    public PtgOperand(int rank, string text)
    {
        Rank = rank;
        Text = text;
    }
}

static class Ptg
{
    const string FUNC_TABLE = "0,COUNT,0,30;1,IF,2,3;2,ISNA,1,1;3,ISERROR,1,1;4,SUM,0,30;5,AVERAGE,1,30;6,MIN,1,30;7,MAX,1,30;8,ROW,0,1;9,COLUMN,0,1;10,NA,0,0;11,NPV,2,30;12,STDEV,1,30;13,DOLLAR,1,2;14,FIXED,2,3;15,SIN,1,1;16,COS,1,1;17,TAN,1,1;18,ATAN,1,1;19,PI,0,0;20,SQRT,1,1;21,EXP,1,1;22,LN,1,1;23,LOG10,1,1;24,ABS,1,1;25,INT,1,1;26,SIGN,1,1;27,ROUND,2,2;28,LOOKUP,2,3;29,INDEX,2,4;30,REPT,2,2;31,MID,3,3;32,LEN,1,1;33,VALUE,1,1;34,TRUE,0,0;35,FALSE,0,0;36,AND,1,30;37,OR,1,30;38,NOT,1,1;39,MOD,2,2;40,DCOUNT,3,3;41,DSUM,3,3;42,DAVERAGE,3,3;43,DMIN,3,3;44,DMAX,3,3;45,DSTDEV,3,3;46,VAR,1,30;47,DVAR,3,3;48,TEXT,2,2;49,LINEST,1,4;50,TREND,1,4;51,LOGEST,1,4;52,GROWTH,1,4;56,PV,3,5;57,FV,3,5;58,NPER,3,5;59,PMT,3,5;60,RATE,3,6;61,MIRR,3,3;62,IRR,1,2;63,RAND,0,0;64,MATCH,2,3;65,DATE,3,3;66,TIME,3,3;67,DAY,1,1;68,MONTH,1,1;69,YEAR,1,1;70,WEEKDAY,1,2;71,HOUR,1,1;72,MINUTE,1,1;73,SECOND,1,1;74,NOW,0,0;75,AREAS,1,1;76,ROWS,1,1;77,COLUMNS,1,1;78,OFFSET,3,5;82,SEARCH,2,3;83,TRANSPOSE,1,1;86,TYPE,1,1;92,SERIESSUM,4,4;97,ATAN2,2,2;98,ASIN,1,1;99,ACOS,1,1;100,CHOOSE,2,30;101,HLOOKUP,3,4;102,VLOOKUP,3,4;105,ISREF,1,1;109,LOG,1,2;111,CHAR,1,1;112,LOWER,1,1;113,UPPER,1,1;114,PROPER,1,1;115,LEFT,1,2;116,RIGHT,1,2;117,EXACT,2,2;118,TRIM,1,1;119,REPLACE,4,4;120,SUBSTITUTE,3,4;121,CODE,1,1;124,FIND,2,3;125,CELL,1,2;126,ISERR,1,1;127,ISTEXT,1,1;128,ISNUMBER,1,1;129,ISBLANK,1,1;130,T,1,1;131,N,1,1;140,DATEVALUE,1,1;141,TIMEVALUE,1,1;142,SLN,3,3;143,SYD,4,4;144,DDB,4,5;148,INDIRECT,1,2;162,CLEAN,1,1;163,MDETERM,1,1;164,MINVERSE,1,1;165,MMULT,2,2;167,IPMT,4,6;168,PPMT,4,6;169,COUNTA,0,30;183,PRODUCT,0,30;184,FACT,1,1;189,DPRODUCT,3,3;190,ISNONTEXT,1,1;193,STDEVP,1,30;194,VARP,1,30;195,DSTDEVP,3,3;196,DVARP,3,3;197,TRUNC,1,2;198,ISLOGICAL,1,1;199,DCOUNTA,3,3;204,USDOLLAR,1,2;205,FINDB,2,3;206,SEARCHB,2,3;207,REPLACEB,4,4;208,LEFTB,1,2;209,RIGHTB,1,2;210,MIDB,3,3;211,LENB,1,1;212,ROUNDUP,2,2;213,ROUNDDOWN,2,2;214,ASC,1,1;215,DBCS,1,1;216,RANK,2,3;219,ADDRESS,2,5;220,DAYS360,2,3;221,TODAY,0,0;222,VDB,5,7;227,MEDIAN,1,30;228,SUMPRODUCT,1,30;229,SINH,1,1;230,COSH,1,1;231,TANH,1,1;232,ASINH,1,1;233,ACOSH,1,1;234,ATANH,1,1;235,DGET,3,3;244,INFO,1,1;247,DB,4,5;252,FREQUENCY,2,2;261,ERROR.TYPE,1,1;269,AVEDEV,1,30;270,BETADIST,3,5;271,GAMMALN,1,1;272,BETAINV,3,5;273,BINOMDIST,4,4;274,CHIDIST,2,2;275,CHIINV,2,2;276,COMBIN,2,2;277,CONFIDENCE,3,3;278,CRITBINOM,3,3;279,EVEN,1,1;280,EXPONDIST,3,3;281,FDIST,3,3;282,FINV,3,3;283,FISHER,1,1;284,FISHERINV,1,1;285,FLOOR,2,2;286,GAMMADIST,4,4;287,GAMMAINV,3,3;288,CEILING,2,2;289,HYPGEOMDIST,4,4;290,LOGNORMDIST,3,3;291,LOGINV,3,3;292,NEGBINOMDIST,3,3;293,NORMDIST,4,4;294,NORMSDIST,1,1;295,NORMINV,3,3;296,NORMSINV,1,1;297,STANDARDIZE,3,3;298,ODD,1,1;299,PERMUT,2,2;300,POISSON,3,3;301,TDIST,3,3;302,WEIBULL,4,4;303,SUMXMY2,2,2;304,SUMX2MY2,2,2;305,SUMX2PY2,2,2;306,CHITEST,2,2;307,CORREL,2,2;308,COVAR,2,2;309,FORECAST,3,3;310,FTEST,2,2;311,INTERCEPT,2,2;312,PEARSON,2,2;313,RSQ,2,2;314,STEYX,2,2;315,SLOPE,2,2;316,TTEST,4,4;317,PROB,3,4;318,DEVSQ,1,30;319,GEOMEAN,1,30;320,HARMEAN,1,30;321,SUMSQ,0,30;322,KURT,1,30;323,SKEW,1,30;324,ZTEST,2,3;325,LARGE,2,2;326,SMALL,2,2;327,QUARTILE,2,2;328,PERCENTILE,2,2;329,PERCENTRANK,2,3;330,MODE,1,30;331,TRIMMEAN,2,2;332,TINV,2,2;336,CONCATENATE,0,30;337,POWER,2,2;342,RADIANS,1,1;343,DEGREES,1,1;344,SUBTOTAL,2,30;345,SUMIF,2,3;346,COUNTIF,2,2;347,COUNTBLANK,1,1;350,ISPMT,4,4;351,DATEDIF,3,3;352,DATESTRING,1,1;353,NUMBERSTRING,2,2;354,ROMAN,1,2;358,GETPIVOTDATA,2,2;359,HYPERLINK,1,2;360,PHONETIC,1,1;361,AVERAGEA,1,30;362,MAXA,1,30;363,MINA,1,30;364,STDEVPA,1,30;365,VARPA,1,30;366,STDEVA,1,30;367,VARA,1,30;368,BAHTTEXT,1,1;369,THAIDAYOFWEEK,1,1;370,THAIDIGIT,1,1;371,THAIMONTHOFYEAR,1,1;372,THAINUMSOUND,1,1;373,THAINUMSTRING,1,1;374,THAISTRINGLENGTH,1,1;375,ISTHAIDIGIT,1,1;376,ROUNDBAHTDOWN,1,1;377,ROUNDBAHTUP,1,1;378,THAIYEAR,1,1;379,RTD,2,5";

    internal class FuncInfo
    {
        public string Name;
        public int Min;
        public int Max;
    }

    internal static Dictionary<int, FuncInfo> funcs;

    static Ptg()
    {
        funcs = new Dictionary<int, FuncInfo>();
        string[] parts = FUNC_TABLE.Split(';');
        for (int i = 0; i < parts.Length; i++)
        {
            string[] f = parts[i].Split(',');
            FuncInfo fi = new FuncInfo();
            fi.Name = f[1];
            fi.Min = int.Parse(f[2], CultureInfo.InvariantCulture);
            fi.Max = int.Parse(f[3], CultureInfo.InvariantCulture);
            funcs[int.Parse(f[0], CultureInfo.InvariantCulture)] = fi;
        }
    }

    internal static FuncInfo GetFunc(int id)
    {
        FuncInfo fi;
        if (funcs.TryGetValue(id, out fi)) return fi;
        return null;
    }

    // ------------------------------------------------------------ public

    public static string Decompile(byte[] data, int len, FormulaContext ctx)
    {
        List<PtgOperand> stack = new List<PtgOperand>();
        int pos = 0;

        while (pos < len)
        {
            int op = data[pos];
            int opcode = op & 0x1F;
            int optype = (op & 0x60) >> 5;
            int opx = optype != 0 ? opcode + 32 : opcode;

            switch (opx)
            {
                case 1: // PtgExp (shared formula)
                    {
                        int row = ReadU16(data, pos + 1);
                        int col = ReadU16(data, pos + 3);
                        string text = "#SHARED";
                        if (ctx.SharedFormula != null)
                        {
                            byte[] tokens = ctx.SharedFormula(row, col);
                            if (tokens != null && tokens.Length > 0)
                            {
                                FormulaContext sub = new FormulaContext();
                                sub.SheetNames = ctx.SheetNames;
                                sub.Names = ctx.Names;
                                sub.ExternSheet = ctx.ExternSheet;
                                sub.RelDelta = true;
                                sub.BaseRow = row;
                                sub.BaseCol = col;
                                sub.RowOffset = ctx.CellRow - row;
                                sub.ColOffset = ctx.CellCol - col;
                                sub.SharedFormula = ctx.SharedFormula;
                                sub.CellRow = ctx.CellRow;
                                sub.CellCol = ctx.CellCol;
                                text = Decompile(tokens, tokens.Length, sub);
                            }
                        }
                        stack.Add(new PtgOperand(90, text));
                        pos += 5;
                    }
                    break;

                case 2: // PtgTbl (array formula master)
                    stack.Add(new PtgOperand(90, "#ARRAY"));
                    pos += 5;
                    break;

                case 3: BinOp(stack, "+", 30); pos += 1; break;
                case 4: BinOp(stack, "-", 30); pos += 1; break;
                case 5: BinOp(stack, "*", 40); pos += 1; break;
                case 6: BinOp(stack, "/", 40); pos += 1; break;
                case 7: BinOp(stack, "^", 50); pos += 1; break;
                case 8: BinOp(stack, "&", 20); pos += 1; break;
                case 9: BinOp(stack, "<", 10); pos += 1; break;
                case 10: BinOp(stack, "<=", 10); pos += 1; break;
                case 11: BinOp(stack, "=", 10); pos += 1; break;
                case 12: BinOp(stack, ">=", 10); pos += 1; break;
                case 13: BinOp(stack, ">", 10); pos += 1; break;
                case 14: BinOp(stack, "<>", 10); pos += 1; break;
                case 15: BinOp(stack, " ", 80); pos += 1; break;
                case 16: BinOp(stack, ",", 80); pos += 1; break;
                case 17: BinOp(stack, ":", 80); pos += 1; break;

                case 18: Unary(stack, "+", "", 70); pos += 1; break;
                case 19: Unary(stack, "-", "", 70); pos += 1; break;
                case 20: Unary(stack, "", "%", 60); pos += 1; break;

                case 21: // PtgParen
                    {
                        if (stack.Count > 0)
                        {
                            PtgOperand a = stack[stack.Count - 1];
                            stack[stack.Count - 1] = new PtgOperand(90, "(" + a.Text + ")");
                        }
                        pos += 1;
                    }
                    break;

                case 22: stack.Add(new PtgOperand(90, "")); pos += 1; break; // PtgMissArg

                case 23: // PtgStr
                    {
                        int cch = data[pos + 1];
                        int flags = data[pos + 2];
                        int p = pos + 3;
                        int rt = 0, ph = 0;
                        if ((flags & 0x08) != 0) { rt = ReadU16(data, p); p += 2; }
                        if ((flags & 0x04) != 0) { ph = (int)ReadU32(data, p); p += 4; }
                        string s;
                        if ((flags & 0x01) != 0)
                        {
                            s = Encoding.Unicode.GetString(data, p, cch * 2);
                            p += cch * 2;
                        }
                        else
                        {
                            s = Encoding.GetEncoding(28591).GetString(data, p, cch);
                            p += cch;
                        }
                        p += rt * 4 + ph;
                        stack.Add(new PtgOperand(90, "\"" + s.Replace("\"", "\"\"") + "\""));
                        pos = p;
                    }
                    break;

                case 24: stack.Add(new PtgOperand(90, "#EXT")); pos += 1; break; // PtgExtended
                case 25: // PtgAttr
                    {
                        int subop = data[pos + 1];
                        int nc = ReadU16(data, pos + 2);
                        int size = 4;
                        if (subop == 0x04) size = nc * 2 + 6; // tAttrChoose
                        else if (subop == 0x10) // tAttrSum
                        {
                            if (stack.Count > 0)
                            {
                                PtgOperand a = stack[stack.Count - 1];
                                stack[stack.Count - 1] = new PtgOperand(90, "SUM(" + a.Text + ")");
                            }
                        }
                        pos += size;
                    }
                    break;

                case 26: stack.Add(new PtgOperand(90, "#SHEET")); pos += 1; break;
                case 27: stack.Add(new PtgOperand(90, "#ENDSHEET")); pos += 1; break;

                case 28: // PtgErr
                    stack.Add(new PtgOperand(90, ErrText(data[pos + 1])));
                    pos += 2;
                    break;

                case 29: // PtgBool
                    stack.Add(new PtgOperand(90, data[pos + 1] != 0 ? "TRUE" : "FALSE"));
                    pos += 2;
                    break;

                case 30: // PtgInt
                    stack.Add(new PtgOperand(90, ReadU16(data, pos + 1).ToString(CultureInfo.InvariantCulture)));
                    pos += 3;
                    break;

                case 31: // PtgNum
                    stack.Add(new PtgOperand(90, NumStr(BitConverter.ToDouble(data, pos + 1))));
                    pos += 9;
                    break;

                case 32: // PtgArray
                    stack.Add(new PtgOperand(90, "#ARRAY"));
                    pos += 9;
                    break;

                case 33: // PtgFunc
                    {
                        int id = ReadU16(data, pos + 1);
                        FuncInfo fi;
                        string name = funcs.TryGetValue(id, out fi) ? fi.Name : "FUNC" + id;
                        int nargs = fi != null ? fi.Min : 0;
                        BuildCall(stack, name, nargs);
                        pos += 3;
                    }
                    break;

                case 34: // PtgFuncVar
                    {
                        int cp = data[pos + 1];
                        int id = ReadU16(data, pos + 2) & 0x7FFF;
                        int nargs = cp & 0x7F;
                        FuncInfo fi;
                        if (id == 255 && nargs > 0 && stack.Count >= nargs)
                        {
                            int start = stack.Count - nargs;
                            string fname = stack[start].Text;
                            stack.RemoveAt(start);
                            BuildCall(stack, fname, nargs - 1);
                        }
                        else
                        {
                            string name = funcs.TryGetValue(id, out fi) ? fi.Name : "FUNC" + id;
                            BuildCall(stack, name, nargs);
                        }
                        pos += 4;
                    }
                    break;

                case 35: // PtgName
                    {
                        int idx = (int)ReadU32(data, pos + 1) - 1;
                        string text = "#NAME";
                        if (ctx.Names != null && idx >= 0 && idx < ctx.Names.Count)
                        {
                            DefinedName dn = ctx.Names[idx];
                            text = dn.Name;
                            if (dn.LocalSheetId >= 0 && ctx.SheetNames != null && dn.LocalSheetId < ctx.SheetNames.Length)
                                text = QuoteSheet(ctx.SheetNames[dn.LocalSheetId]) + "!" + text;
                        }
                        stack.Add(new PtgOperand(90, text));
                        pos += 5;
                    }
                    break;

                case 36: // PtgRef
                    {
                        int rowval = ReadU16(data, pos + 1);
                        int colval = ReadU16(data, pos + 3);
                        bool rowRel = ((colval >> 15) & 1) != 0;
                        bool colRel = ((colval >> 14) & 1) != 0;
                        int row = rowval;
                        int col = colval & 0xFF;
                        Adjust(ref row, ref col, rowRel, colRel, ctx);
                        stack.Add(new PtgOperand(90, RefText(row, col, rowRel, colRel)));
                        pos += 5;
                    }
                    break;

                case 37: // PtgArea
                    {
                        int row1 = ReadU16(data, pos + 1);
                        int row2 = ReadU16(data, pos + 3);
                        int col1 = ReadU16(data, pos + 5);
                        int col2 = ReadU16(data, pos + 7);
                        bool r1rel = ((col1 >> 15) & 1) != 0;
                        bool c1rel = ((col1 >> 14) & 1) != 0;
                        bool r2rel = ((col2 >> 15) & 1) != 0;
                        bool c2rel = ((col2 >> 14) & 1) != 0;
                        int r1 = row1, c1 = col1 & 0xFF, r2 = row2, c2 = col2 & 0xFF;
                        Adjust(ref r1, ref c1, r1rel, c1rel, ctx);
                        Adjust(ref r2, ref c2, r2rel, c2rel, ctx);
                        string text = RefText(r1, c1, r1rel, c1rel);
                        if (r1 != r2 || c1 != c2) text += ":" + RefText(r2, c2, r2rel, c2rel);
                        stack.Add(new PtgOperand(90, text));
                        pos += 9;
                    }
                    break;

                case 38: pos += 7; break; // PtgMemArea
                case 39: pos += 7; break; // PtgMemErr
                case 40: pos += 7; break; // PtgMemNoMem
                case 41: pos += 3; break; // PtgMemFunc

                case 42: stack.Add(new PtgOperand(90, "#REF!")); pos += 5; break; // PtgRefErr
                case 43: stack.Add(new PtgOperand(90, "#REF!")); pos += 9; break; // PtgAreaErr

                case 44: // PtgRefN
                    {
                        int rowval = ReadU16(data, pos + 1);
                        int colval = ReadU16(data, pos + 3);
                        bool rowRel = ((colval >> 15) & 1) != 0;
                        bool colRel = ((colval >> 14) & 1) != 0;
                        int row = rowval;
                        int col = colval & 0xFF;
                        AdjustN(ref row, ref col, ctx);
                        stack.Add(new PtgOperand(90, RefText(row, col, rowRel, colRel)));
                        pos += 5;
                    }
                    break;

                case 45: // PtgAreaN
                    {
                        int row1 = ReadU16(data, pos + 1);
                        int row2 = ReadU16(data, pos + 3);
                        int col1 = ReadU16(data, pos + 5);
                        int col2 = ReadU16(data, pos + 7);
                        bool r1rel = ((col1 >> 15) & 1) != 0;
                        bool c1rel = ((col1 >> 14) & 1) != 0;
                        bool r2rel = ((col2 >> 15) & 1) != 0;
                        bool c2rel = ((col2 >> 14) & 1) != 0;
                        int r1 = row1, c1 = col1 & 0xFF, r2 = row2, c2 = col2 & 0xFF;
                        AdjustN(ref r1, ref c1, ctx);
                        AdjustN(ref r2, ref c2, ctx);
                        string text = RefText(r1, c1, r1rel, c1rel);
                        if (r1 != r2 || c1 != c2) text += ":" + RefText(r2, c2, r2rel, c2rel);
                        stack.Add(new PtgOperand(90, text));
                        pos += 9;
                    }
                    break;

                case 46: pos += 3; break; // PtgMemAreaN
                case 47: pos += 3; break; // PtgMemNoMemN

                case 56: // PtgFuncCE
                    {
                        int id = ReadU16(data, pos + 1);
                        FuncInfo fi;
                        string name = funcs.TryGetValue(id, out fi) ? fi.Name : "FUNC" + id;
                        BuildCall(stack, name, fi != null ? fi.Min : 0);
                        pos += 3;
                    }
                    break;

                case 57: // PtgNameX
                    {
                        int idx = (int)ReadU32(data, pos + 3) - 1;
                        string text = "#NAME";
                        if (ctx.Names != null && idx >= 0 && idx < ctx.Names.Count) text = ctx.Names[idx].Name;
                        stack.Add(new PtgOperand(90, text));
                        pos += 7;
                    }
                    break;

                case 58: // PtgRef3d
                    {
                        int ixti = ReadU16(data, pos + 1);
                        int rowval = ReadU16(data, pos + 3);
                        int colval = ReadU16(data, pos + 5);
                        bool rowRel = ((colval >> 15) & 1) != 0;
                        bool colRel = ((colval >> 14) & 1) != 0;
                        int row = rowval;
                        int col = colval & 0xFF;
                        Adjust(ref row, ref col, rowRel, colRel, ctx);
                        string sheet = ctx.ExternSheet != null ? ctx.ExternSheet(ixti) : null;
                        string text = (sheet != null && sheet.Length > 0 ? QuoteSheet(sheet) + "!" : "") + RefText(row, col, rowRel, colRel);
                        stack.Add(new PtgOperand(90, text));
                        pos += 7;
                    }
                    break;

                case 59: // PtgArea3d
                    {
                        int ixti = ReadU16(data, pos + 1);
                        int row1 = ReadU16(data, pos + 3);
                        int row2 = ReadU16(data, pos + 5);
                        int col1 = ReadU16(data, pos + 7);
                        int col2 = ReadU16(data, pos + 9);
                        bool r1rel = ((col1 >> 15) & 1) != 0;
                        bool c1rel = ((col1 >> 14) & 1) != 0;
                        bool r2rel = ((col2 >> 15) & 1) != 0;
                        bool c2rel = ((col2 >> 14) & 1) != 0;
                        int r1 = row1, c1 = col1 & 0xFF, r2 = row2, c2 = col2 & 0xFF;
                        Adjust(ref r1, ref c1, r1rel, c1rel, ctx);
                        Adjust(ref r2, ref c2, r2rel, c2rel, ctx);
                        string sheet = ctx.ExternSheet != null ? ctx.ExternSheet(ixti) : null;
                        string range = RefText(r1, c1, r1rel, c1rel);
                        if (r1 != r2 || c1 != c2) range += ":" + RefText(r2, c2, r2rel, c2rel);
                        string text = (sheet != null && sheet.Length > 0 ? QuoteSheet(sheet) + "!" : "") + range;
                        stack.Add(new PtgOperand(90, text));
                        pos += 11;
                    }
                    break;

                case 60: stack.Add(new PtgOperand(90, "#REF!")); pos += 7; break; // PtgRefErr3d
                case 61: stack.Add(new PtgOperand(90, "#REF!")); pos += 11; break; // PtgAreaErr3d

                default:
                    pos = len;
                    break;
            }
        }

        return stack.Count > 0 ? stack[stack.Count - 1].Text : "";
    }

    // ------------------------------------------------------------ helpers

    internal static int ReadU16(byte[] d, int p)
    {
        return d[p] | (d[p + 1] << 8);
    }

    internal static uint ReadU32(byte[] d, int p)
    {
        return (uint)(d[p] | (d[p + 1] << 8) | (d[p + 2] << 16) | (d[p + 3] << 24));
    }

    internal static void BinOp(List<PtgOperand> stack, string sym, int rank)
    {
        if (stack.Count < 2) return;
        PtgOperand b = stack[stack.Count - 1];
        stack.RemoveAt(stack.Count - 1);
        PtgOperand a = stack[stack.Count - 1];
        stack.RemoveAt(stack.Count - 1);
        string at = a.Rank < rank ? "(" + a.Text + ")" : a.Text;
        string bt = b.Rank < rank ? "(" + b.Text + ")" : b.Text;
        stack.Add(new PtgOperand(rank, at + sym + bt));
    }

    internal static void Unary(List<PtgOperand> stack, string pre, string post, int rank)
    {
        if (stack.Count < 1) return;
        PtgOperand a = stack[stack.Count - 1];
        stack.RemoveAt(stack.Count - 1);
        string t = a.Rank < rank ? "(" + a.Text + ")" : a.Text;
        stack.Add(new PtgOperand(rank, pre + t + post));
    }

    internal static void BuildCall(List<PtgOperand> stack, string name, int nargs)
    {
        if (nargs > stack.Count) nargs = stack.Count;
        string[] args = new string[nargs];
        for (int i = nargs - 1; i >= 0; i--)
        {
            args[i] = stack[stack.Count - 1].Text;
            stack.RemoveAt(stack.Count - 1);
        }
        stack.Add(new PtgOperand(90, name + "(" + string.Join(",", args) + ")"));
    }

    static void Adjust(ref int row, ref int col, bool rowRel, bool colRel, FormulaContext ctx)
    {
        if (!ctx.RelDelta) return;
        if (rowRel && row >= 32768) row -= 65536;
        if (colRel && col >= 128) col -= 256;
        row += ctx.BaseRow + ctx.RowOffset;
        col += ctx.BaseCol + ctx.ColOffset;
    }

    static void AdjustN(ref int row, ref int col, FormulaContext ctx)
    {
        if (row >= 32768) row -= 65536;
        if (col >= 128) col -= 256;
        row += ctx.BaseRow + ctx.RowOffset;
        col += ctx.BaseCol + ctx.ColOffset;
    }

    internal static string RefText(int row, int col, bool rowRel, bool colRel)
    {
        string c = A1.IndexToCol(col + 1);
        string r = (row + 1).ToString(CultureInfo.InvariantCulture);
        if (!colRel) c = "$" + c;
        if (!rowRel) r = "$" + r;
        return c + r;
    }

    internal static string QuoteSheet(string s)
    {
        if (s == null) return "";
        if (s.IndexOf('\'') >= 0) return "'" + s.Replace("'", "''") + "'";
        if (s.IndexOf(' ') >= 0) return "'" + s + "'";
        return s;
    }

    internal static string NumStr(double d)
    {
        string s = d.ToString("G15", CultureInfo.InvariantCulture);
        if (s.EndsWith(".0")) s = s.Substring(0, s.Length - 2);
        return s;
    }

    internal static string ErrText(int code)
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
}
