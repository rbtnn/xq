using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

// BIFF12 (.xlsb) formula token decompiler. Token payload sizes follow
// MS-XLSB 2.5.98 / calamine's xlsb parser; function IDs are the same as BIFF8.
static class Ptg12
{
    public static string Decompile(byte[] t, int len, FormulaContext ctx)
    {
        List<PtgOperand> stack = new List<PtgOperand>();
        int p = 0;

        while (p < len)
        {
            int ptg = t[p];
            p++;

            switch (ptg)
            {
                // ---- references ------------------------------------------------
                case 0x24: case 0x44: case 0x64: // PtgRef
                    {
                        int row = (int)Ptg.ReadU32(t, p);
                        int colraw = Ptg.ReadU16(t, p + 4);
                        p += 6;
                        int col = colraw & 0x3FFF;
                        bool colRel = (colraw & 0x8000) != 0;
                        bool rowRel = (colraw & 0x4000) != 0;
                        if (ctx.RelDelta)
                        {
                            if (rowRel) row += ctx.BaseRow + ctx.RowOffset;
                            if (colRel) col += ctx.BaseCol + ctx.ColOffset;
                        }
                        stack.Add(new PtgOperand(90, Ptg.RefText(row, col, rowRel, colRel)));
                    }
                    break;

                case 0x2C: case 0x4C: case 0x6C: // PtgRefN
                    {
                        int row = (int)Ptg.ReadU32(t, p);
                        int colraw = Ptg.ReadU16(t, p + 4);
                        p += 6;
                        int col = colraw & 0x3FFF;
                        row += ctx.BaseRow + ctx.RowOffset;
                        col += ctx.BaseCol + ctx.ColOffset;
                        stack.Add(new PtgOperand(90, Ptg.RefText(row, col, true, true)));
                    }
                    break;

                case 0x2D: case 0x4D: case 0x6D: // PtgAreaN
                    {
                        int r1 = (int)Ptg.ReadU32(t, p);
                        int r2 = (int)Ptg.ReadU32(t, p + 4);
                        int c1 = Ptg.ReadU16(t, p + 8) & 0x3FFF;
                        int c2 = Ptg.ReadU16(t, p + 10) & 0x3FFF;
                        p += 12;
                        r1 += ctx.BaseRow + ctx.RowOffset;
                        r2 += ctx.BaseRow + ctx.RowOffset;
                        c1 += ctx.BaseCol + ctx.ColOffset;
                        c2 += ctx.BaseCol + ctx.ColOffset;
                        string text = Ptg.RefText(r1, c1, true, true);
                        if (r1 != r2 || c1 != c2)
                            text += ":" + Ptg.RefText(r2, c2, true, true);
                        stack.Add(new PtgOperand(90, text));
                    }
                    break;

                case 0x25: case 0x45: case 0x65: // PtgArea
                    {
                        int r1 = (int)Ptg.ReadU32(t, p);
                        int r2 = (int)Ptg.ReadU32(t, p + 4);
                        int c1raw = Ptg.ReadU16(t, p + 8);
                        int c2raw = Ptg.ReadU16(t, p + 10);
                        p += 12;
                        int c1 = c1raw & 0x3FFF;
                        int c2 = c2raw & 0x3FFF;
                        bool c1rel = (c1raw & 0x8000) != 0;
                        bool c2rel = (c2raw & 0x8000) != 0;
                        bool r1rel = (c1raw & 0x4000) != 0;
                        bool r2rel = (c2raw & 0x4000) != 0;
                        if (ctx.RelDelta)
                        {
                            if (r1rel) r1 += ctx.BaseRow + ctx.RowOffset;
                            if (r2rel) r2 += ctx.BaseRow + ctx.RowOffset;
                            if (c1rel) c1 += ctx.BaseCol + ctx.ColOffset;
                            if (c2rel) c2 += ctx.BaseCol + ctx.ColOffset;
                        }
                        string text = Ptg.RefText(r1, c1, r1rel, c1rel);
                        if (r1 != r2 || c1 != c2)
                            text += ":" + Ptg.RefText(r2, c2, r2rel, c2rel);
                        stack.Add(new PtgOperand(90, text));
                    }
                    break;

                case 0x3A: case 0x5A: case 0x7A: // PtgRef3d
                    {
                        int ixti = Ptg.ReadU16(t, p);
                        int row = (int)Ptg.ReadU32(t, p + 2);
                        int colraw = Ptg.ReadU16(t, p + 6);
                        p += 8;
                        int col = colraw & 0x3FFF;
                        bool colRel = (colraw & 0x8000) != 0;
                        bool rowRel = (colraw & 0x4000) != 0;
                        if (ctx.RelDelta)
                        {
                            if (rowRel) row += ctx.BaseRow + ctx.RowOffset;
                            if (colRel) col += ctx.BaseCol + ctx.ColOffset;
                        }
                        string sheet = ctx.ExternSheet != null ? ctx.ExternSheet(ixti) : null;
                        string text = (sheet != null && sheet.Length > 0 ? Ptg.QuoteSheet(sheet) + "!" : "")
                            + Ptg.RefText(row, col, rowRel, colRel);
                        stack.Add(new PtgOperand(90, text));
                    }
                    break;

                case 0x3B: case 0x5B: case 0x7B: // PtgArea3d
                    {
                        int ixti = Ptg.ReadU16(t, p);
                        int r1 = (int)Ptg.ReadU32(t, p + 2);
                        int r2 = (int)Ptg.ReadU32(t, p + 6);
                        int c1raw = Ptg.ReadU16(t, p + 10);
                        int c2raw = Ptg.ReadU16(t, p + 12);
                        p += 14;
                        int c1 = c1raw & 0x3FFF;
                        int c2 = c2raw & 0x3FFF;
                        bool c1rel = (c1raw & 0x8000) != 0;
                        bool c2rel = (c2raw & 0x8000) != 0;
                        bool r1rel = (c1raw & 0x4000) != 0;
                        bool r2rel = (c2raw & 0x4000) != 0;
                        if (ctx.RelDelta)
                        {
                            if (r1rel) r1 += ctx.BaseRow + ctx.RowOffset;
                            if (r2rel) r2 += ctx.BaseRow + ctx.RowOffset;
                            if (c1rel) c1 += ctx.BaseCol + ctx.ColOffset;
                            if (c2rel) c2 += ctx.BaseCol + ctx.ColOffset;
                        }
                        string sheet = ctx.ExternSheet != null ? ctx.ExternSheet(ixti) : null;
                        string range = Ptg.RefText(r1, c1, r1rel, c1rel);
                        if (r1 != r2 || c1 != c2)
                            range += ":" + Ptg.RefText(r2, c2, r2rel, c2rel);
                        string text = (sheet != null && sheet.Length > 0 ? Ptg.QuoteSheet(sheet) + "!" : "") + range;
                        stack.Add(new PtgOperand(90, text));
                    }
                    break;

                case 0x2A: case 0x4A: case 0x6A: // PtgRefErr
                    stack.Add(new PtgOperand(90, "#REF!"));
                    p += 6;
                    break;
                case 0x2B: case 0x4B: case 0x6B: // PtgAreaErr
                    stack.Add(new PtgOperand(90, "#REF!"));
                    p += 12;
                    break;
                case 0x3C: case 0x5C: case 0x7C: // PtgRefErr3d
                    {
                        int ixti = Ptg.ReadU16(t, p);
                        p += 8;
                        string sheet = ctx.ExternSheet != null ? ctx.ExternSheet(ixti) : null;
                        stack.Add(new PtgOperand(90, (sheet != null && sheet.Length > 0 ? Ptg.QuoteSheet(sheet) + "!" : "") + "#REF!"));
                    }
                    break;
                case 0x3D: case 0x5D: case 0x7D: // PtgAreaErr3d
                    {
                        int ixti = Ptg.ReadU16(t, p);
                        p += 14;
                        string sheet = ctx.ExternSheet != null ? ctx.ExternSheet(ixti) : null;
                        stack.Add(new PtgOperand(90, (sheet != null && sheet.Length > 0 ? Ptg.QuoteSheet(sheet) + "!" : "") + "#REF!"));
                    }
                    break;

                case 0x23: case 0x43: case 0x63: // PtgName
                    {
                        int index = (int)Ptg.ReadU32(t, p) - 1;
                        p += 4;
                        string name = "#NAME?";
                        if (ctx.Names != null && index >= 0 && index < ctx.Names.Count)
                            name = ctx.Names[index].Name;
                        stack.Add(new PtgOperand(90, name));
                    }
                    break;

                case 0x39: case 0x59: case 0x79: // PtgNameX
                    stack.Add(new PtgOperand(90, "#NAME?"));
                    p += 6;
                    break;

                case 0x01: // PtgExp (shared/array formula)
                    {
                        int masterRow = (int)Ptg.ReadU32(t, p);
                        p += 4;
                        string text = "#SHARED";
                        if (ctx.SharedFormulaEx != null)
                        {
                            int baseRow, baseCol;
                            byte[] shared = ctx.SharedFormulaEx(masterRow, ctx.CellCol, out baseRow, out baseCol);
                            if (shared != null && shared.Length > 0)
                            {
                                FormulaContext sub = new FormulaContext();
                                sub.SheetNames = ctx.SheetNames;
                                sub.Names = ctx.Names;
                                sub.ExternSheet = ctx.ExternSheet;
                                sub.RelDelta = true;
                                sub.BaseRow = baseRow;
                                sub.BaseCol = baseCol;
                                sub.RowOffset = ctx.CellRow - baseRow;
                                sub.ColOffset = ctx.CellCol - baseCol;
                                sub.SharedFormulaEx = ctx.SharedFormulaEx;
                                sub.CellRow = ctx.CellRow;
                                sub.CellCol = ctx.CellCol;
                                text = Decompile(shared, shared.Length, sub);
                            }
                        }
                        stack.Add(new PtgOperand(90, text));
                    }
                    break;

                case 0x02: // PtgTbl (array formula master)
                    stack.Add(new PtgOperand(90, "#ARRAY"));
                    p += 4;
                    break;

                // ---- operators --------------------------------------------------
                case 0x03: Ptg.BinOp(stack, "+", 30); break;
                case 0x04: Ptg.BinOp(stack, "-", 30); break;
                case 0x05: Ptg.BinOp(stack, "*", 40); break;
                case 0x06: Ptg.BinOp(stack, "/", 40); break;
                case 0x07: Ptg.BinOp(stack, "^", 50); break;
                case 0x08: Ptg.BinOp(stack, "&", 20); break;
                case 0x09: Ptg.BinOp(stack, "<", 10); break;
                case 0x0A: Ptg.BinOp(stack, "<=", 10); break;
                case 0x0B: Ptg.BinOp(stack, "=", 10); break;
                case 0x0C: Ptg.BinOp(stack, ">", 10); break;
                case 0x0D: Ptg.BinOp(stack, ">=", 10); break;
                case 0x0E: Ptg.BinOp(stack, "<>", 10); break;
                case 0x0F: Ptg.BinOp(stack, " ", 80); break;
                case 0x10: Ptg.BinOp(stack, ",", 80); break;
                case 0x11: Ptg.BinOp(stack, ":", 80); break;

                case 0x12: Ptg.Unary(stack, "+", "", 70); break;
                case 0x13: Ptg.Unary(stack, "-", "", 70); break;
                case 0x14: Ptg.Unary(stack, "", "%", 60); break;
                case 0x15: // PtgParen
                    if (stack.Count > 0)
                    {
                        PtgOperand a = stack[stack.Count - 1];
                        stack[stack.Count - 1] = new PtgOperand(90, "(" + a.Text + ")");
                    }
                    break;
                case 0x16: stack.Add(new PtgOperand(90, "")); break; // PtgMissArg

                // ---- literals ---------------------------------------------------
                case 0x17: // PtgStr
                    {
                        int cch = Ptg.ReadU16(t, p);
                        p += 2;
                        string s = Encoding.Unicode.GetString(t, p, cch * 2);
                        p += cch * 2;
                        stack.Add(new PtgOperand(90, "\"" + s.Replace("\"", "\"\"") + "\""));
                    }
                    break;

                case 0x18: // PtgElf
                    {
                        int eptg = t[p];
                        p += 1;
                        if (eptg == 0x19) p += 12;
                        else if (eptg == 0x1D) p += 4;
                    }
                    break;

                case 0x19: // PtgAttr
                    {
                        int eptg = t[p];
                        p += 1;
                        if (eptg == 0x04) p += 10;
                        else if (eptg == 0x10)
                        {
                            p += 2;
                            if (stack.Count > 0)
                            {
                                PtgOperand a = stack[stack.Count - 1];
                                stack[stack.Count - 1] = new PtgOperand(90, "SUM(" + a.Text + ")");
                            }
                        }
                        else p += 2;
                    }
                    break;

                case 0x1C: // PtgErr
                    stack.Add(new PtgOperand(90, Ptg.ErrText(t[p])));
                    p += 1;
                    break;

                case 0x1D: // PtgBool
                    stack.Add(new PtgOperand(90, t[p] != 0 ? "TRUE" : "FALSE"));
                    p += 1;
                    break;

                case 0x1E: // PtgInt
                    stack.Add(new PtgOperand(90, Ptg.ReadU16(t, p).ToString(CultureInfo.InvariantCulture)));
                    p += 2;
                    break;

                case 0x1F: // PtgNum
                    stack.Add(new PtgOperand(90, Ptg.NumStr(BitConverter.ToDouble(t, p))));
                    p += 8;
                    break;

                case 0x20: case 0x40: case 0x60: // PtgArray
                    stack.Add(new PtgOperand(90, "{...}"));
                    p += 14;
                    break;

                // ---- functions --------------------------------------------------
                case 0x21: case 0x41: case 0x61: // PtgFunc
                    {
                        int id = Ptg.ReadU16(t, p);
                        p += 2;
                        Ptg.FuncInfo fi = Ptg.GetFunc(id);
                        string name = fi != null ? fi.Name : "FUNC" + id;
                        int nargs = fi != null ? fi.Min : 0;
                        Ptg.BuildCall(stack, name, nargs);
                    }
                    break;

                case 0x22: case 0x42: case 0x62: // PtgFuncVar
                    {
                        int argc = t[p];
                        int id = Ptg.ReadU16(t, p + 1);
                        p += 3;
                        Ptg.FuncInfo fi = Ptg.GetFunc(id);
                        if (id == 255 && argc > 0 && stack.Count >= argc)
                        {
                            // User-defined / future function: the first argument is the name.
                            int start = stack.Count - argc;
                            string fname = stack[start].Text;
                            stack.RemoveAt(start);
                            Ptg.BuildCall(stack, fname, argc - 1);
                        }
                        else
                        {
                            string name = fi != null ? fi.Name : "FUNC" + id;
                            Ptg.BuildCall(stack, name, argc);
                        }
                    }
                    break;

                case 0x29: case 0x49: case 0x69: // PtgMemArea
                    {
                        int cce = Ptg.ReadU16(t, p);
                        p += 2 + cce;
                    }
                    break;

                default:
                    // Unknown token: stop rather than silently producing garbage.
                    stack.Add(new PtgOperand(90, "#UNKNOWN" + ptg.ToString("X2", CultureInfo.InvariantCulture)));
                    p = len;
                    break;
            }
        }

        if (stack.Count == 1) return stack[0].Text;
        return null;
    }
}
