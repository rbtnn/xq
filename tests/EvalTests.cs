using System;
using System.Collections.Generic;
using System.Globalization;

class EvalTests
{
    static int failures = 0;
    static int checks = 0;

    static void Check(EvalContext ctx, string expr, string expected)
    {
        checks++;
        try
        {
            EvalValue v = ctx.Evaluate(expr);
            string got = v.ToDisplayString();
            if (got != expected)
            {
                failures++;
                Console.WriteLine("FAIL: " + expr + " -> [" + got + "] expected [" + expected + "]");
            }
        }
        catch (Exception ex)
        {
            failures++;
            Console.WriteLine("FAIL: " + expr + " threw " + ex.GetType().Name + ": " + ex.Message);
        }
    }

    static void CheckError(EvalContext ctx, string expr, string expectedError)
    {
        checks++;
        try
        {
            EvalValue v = ctx.Evaluate(expr);
            string got = v.ToDisplayString();
            if (got != expectedError)
            {
                failures++;
                Console.WriteLine("FAIL: " + expr + " -> [" + got + "] expected error [" + expectedError + "]");
            }
        }
        catch (Exception ex)
        {
            failures++;
            Console.WriteLine("FAIL: " + expr + " threw " + ex.GetType().Name + ": " + ex.Message);
        }
    }

    static void CheckParseError(EvalContext ctx, string expr)
    {
        checks++;
        try
        {
            ctx.Evaluate(expr);
            failures++;
            Console.WriteLine("FAIL: " + expr + " should be a parse error");
        }
        catch (EvalParseException)
        {
        }
        catch (Exception ex)
        {
            failures++;
            Console.WriteLine("FAIL: " + expr + " threw " + ex.GetType().Name + ": " + ex.Message);
        }
    }

    static int Main(string[] args)
    {
        string xlsx = args.Length > 0 ? args[0] : "tests\\fixtures\\eval.xlsx";
        string xls = args.Length > 1 ? args[1] : "tests\\fixtures\\eval.xls";
        string xlsb = args.Length > 2 ? args[2] : "tests\\fixtures\\eval.xlsb";

        IBook book = Book.Open(xlsx);
        EvalContext ctx = new EvalContext(book, "Data");

        Check(ctx, "=A1", "1");
        Check(ctx, "=A1+A2", "3");
        Check(ctx, "=SUM(A1:A3)", "6");
        Check(ctx, "=Other!A1", "10");
        Check(ctx, "=SUM(MyRange)", "6");
        Check(ctx, "=IF(A1>0,\"yes\",\"no\")", "yes");
        Check(ctx, "=A1&\"x\"", "1x");
        Check(ctx, "=2^3", "8");
        Check(ctx, "=50%", "0.5");
        Check(ctx, "=ROUND(1.2345,2)", "1.23");
        Check(ctx, "=CONCATENATE(B1,\" \",B2)", "Hello World");
        Check(ctx, "=LEFT(B1,2)", "He");
        Check(ctx, "=LEN(B1)", "5");
        Check(ctx, "=IFERROR(1/0,\"err\")", "err");
        CheckError(ctx, "=NA()", "#N/A");
        CheckError(ctx, "=1/0", "#DIV/0!");
        CheckError(ctx, "=UNKNOWN(1)", "#NAME?");
        Check(ctx, "=TEXT(1234.5,\"#,##0.00\")", "1,234.50");
        Check(ctx, "=YEAR(DATE(2024,1,15))", "2024");
        Check(ctx, "=MONTH(DATE(2024,1,15))", "1");
        Check(ctx, "=DAY(DATE(2024,1,15))", "15");
        Check(ctx, "=VALUE(\"123\")", "123");
        Check(ctx, "=TRIM(\"  a   b  \")", "a b");
        Check(ctx, "=UPPER(\"abc\")", "ABC");
        Check(ctx, "=LOWER(\"ABC\")", "abc");
        Check(ctx, "=MID(\"abcdef\",2,3)", "bcd");
        Check(ctx, "=RIGHT(\"abc\",2)", "bc");
        Check(ctx, "=MOD(7,3)", "1");
        Check(ctx, "=ABS(-3)", "3");
        Check(ctx, "=POWER(2,3)", "8");
        Check(ctx, "=SQRT(9)", "3");
        Check(ctx, "=COUNT(A1:A3)", "3");
        Check(ctx, "=COUNTA(B1:B2)", "2");
        Check(ctx, "=AVERAGE(A1:A3)", "2");
        Check(ctx, "=MIN(A1:A3)", "1");
        Check(ctx, "=MAX(A1:A3)", "3");
        Check(ctx, "=AND(A1>0,A2>0)", "TRUE");
        Check(ctx, "=OR(A1<0,A2>0)", "TRUE");
        Check(ctx, "=NOT(A1>0)", "FALSE");
        Check(ctx, "=ISBLANK(Z99)", "TRUE");
        Check(ctx, "=ISERROR(1/0)", "TRUE");
        Check(ctx, "=ISNA(NA())", "TRUE");
        Check(ctx, "=IF(0,\"a\",\"b\")", "b");
        Check(ctx, "=SUM(A1:A3,10)", "16");
        Check(ctx, "=Other!A1+Data!A2", "12");
        Check(ctx, "='Other'!A1", "10");
        Check(ctx, "=OtherVal", "10");

        Check(ctx, "=A1:B2", "1");
        Check(ctx, "=SUM(A1:B2)", "3");
        Check(ctx, "=COUNT(A1:B2)", "2");
        Check(ctx, "=IF(TRUE,1,2)", "1");
        Check(ctx, "=IF(FALSE,1,2)", "2");
        Check(ctx, "=TRUE()", "TRUE");
        Check(ctx, "=FALSE()", "FALSE");
        Check(ctx, "=ISNUMBER(\"1\")", "FALSE");
        Check(ctx, "=ISNUMBER(A1)", "TRUE");
        Check(ctx, "=ISTEXT(B1)", "TRUE");
        Check(ctx, "=A1=1", "TRUE");
        Check(ctx, "=B1=\"hello\"", "TRUE");
        Check(ctx, "=A1<>2", "TRUE");
        Check(ctx, "=SUM(Other!A1:A2)", "30");
        Check(ctx, "=SUM(MyRange,Other!A1:A2)", "36");
        Check(ctx, "=COUNT(MyRange)", "3");
        Check(ctx, "=SUM(A1:A3)/COUNT(A1:A3)", "2");
        Check(ctx, "=IFERROR(1,2)", "1");
        Check(ctx, "=IFERROR(NA(),\"x\")", "x");
        Check(ctx, "=ROUND(2.5,0)", "3");
        Check(ctx, "=ROUND(-2.5,0)", "-3");
        Check(ctx, "=ROUNDUP(1.21,1)", "1.3");
        Check(ctx, "=ROUNDDOWN(1.29,1)", "1.2");
        Check(ctx, "=MOD(-7,3)", "2");
        Check(ctx, "=TEXT(DATE(2024,1,15),\"yyyy/m/d\")", "2024/1/15");
        Check(ctx, "=TEXT(0.5,\"0%\")", "50%");
        Check(ctx, "=CONCATENATE(\"a\",\"b\")", "ab");
        Check(ctx, "=LEFT(\"abc\")", "a");
        Check(ctx, "=RIGHT(\"abc\")", "c");

        CheckParseError(ctx, "=A1+");
        CheckParseError(ctx, "=SUM(");
        CheckParseError(ctx, "=1+*2");

        EvalContext other = new EvalContext(book, "Other");
        Check(other, "=A1", "10");

        book.Dispose();

        IBook book2 = Book.Open(xls);
        EvalContext ctx2 = new EvalContext(book2, "Data");
        Check(ctx2, "=A1+A2", "3");
        Check(ctx2, "=SUM(A1:A3)", "6");
        Check(ctx2, "=Other!A1", "10");
        Check(ctx2, "=B1", "Hello");
        book2.Dispose();

        IBook book3 = Book.Open(xlsb);
        EvalContext ctx3 = new EvalContext(book3, "Named Ranges");
        Check(ctx3, "=A1", "col1");
        Check(ctx3, "=SUM(col1_)", "15");
        book3.Dispose();

        Console.WriteLine("checks=" + checks + " failures=" + failures);
        return failures == 0 ? 0 : 1;
    }
}
