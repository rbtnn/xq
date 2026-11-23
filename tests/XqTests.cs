using System;
using System.IO;
using System.Text;
using Xunit;

public class XqTests
{
    static string Fixture(string name)
    {
        return Path.Combine(AppContext.BaseDirectory, "fixtures", name);
    }

    static void WithContext(string file, string sheet, Action<EvalContext> action)
    {
        IBook book = Book.Open(Fixture(file));
        try
        {
            action(new EvalContext(book, sheet));
        }
        finally
        {
            book.Dispose();
        }
    }

    static string Eval(string file, string sheet, string expr)
    {
        string result = null;
        WithContext(file, sheet, ctx => result = ctx.Evaluate(expr).ToDisplayString());
        return result;
    }

    [Theory]
    [InlineData("=A1", "1")]
    [InlineData("=A1+A2", "3")]
    [InlineData("=SUM(A1:A3)", "6")]
    [InlineData("=Other!A1", "10")]
    [InlineData("=SUM(MyRange)", "6")]
    [InlineData("=IF(A1>0,\"yes\",\"no\")", "yes")]
    [InlineData("=A1&\"x\"", "1x")]
    [InlineData("=2^3", "8")]
    [InlineData("=50%", "0.5")]
    [InlineData("=ROUND(1.2345,2)", "1.23")]
    [InlineData("=CONCATENATE(B1,\" \",B2)", "Hello World")]
    [InlineData("=LEFT(B1,2)", "He")]
    [InlineData("=LEN(B1)", "5")]
    [InlineData("=IFERROR(1/0,\"err\")", "err")]
    [InlineData("=TEXT(1234.5,\"#,##0.00\")", "1,234.50")]
    [InlineData("=YEAR(DATE(2024,1,15))", "2024")]
    [InlineData("=MONTH(DATE(2024,1,15))", "1")]
    [InlineData("=DAY(DATE(2024,1,15))", "15")]
    [InlineData("=VALUE(\"123\")", "123")]
    [InlineData("=TRIM(\"  a   b  \")", "a b")]
    [InlineData("=UPPER(\"abc\")", "ABC")]
    [InlineData("=LOWER(\"ABC\")", "abc")]
    [InlineData("=MID(\"abcdef\",2,3)", "bcd")]
    [InlineData("=RIGHT(\"abc\",2)", "bc")]
    [InlineData("=MOD(7,3)", "1")]
    [InlineData("=ABS(-3)", "3")]
    [InlineData("=POWER(2,3)", "8")]
    [InlineData("=SQRT(9)", "3")]
    [InlineData("=COUNT(A1:A3)", "3")]
    [InlineData("=COUNTA(B1:B2)", "2")]
    [InlineData("=AVERAGE(A1:A3)", "2")]
    [InlineData("=MIN(A1:A3)", "1")]
    [InlineData("=MAX(A1:A3)", "3")]
    [InlineData("=AND(A1>0,A2>0)", "TRUE")]
    [InlineData("=OR(A1<0,A2>0)", "TRUE")]
    [InlineData("=NOT(A1>0)", "FALSE")]
    [InlineData("=ISBLANK(Z99)", "TRUE")]
    [InlineData("=ISERROR(1/0)", "TRUE")]
    [InlineData("=ISNA(NA())", "TRUE")]
    [InlineData("=IF(0,\"a\",\"b\")", "b")]
    [InlineData("=SUM(A1:A3,10)", "16")]
    [InlineData("=Other!A1+Data!A2", "12")]
    [InlineData("='Other'!A1", "10")]
    [InlineData("=OtherVal", "10")]
    [InlineData("=A1:B2", "1")]
    [InlineData("=SUM(A1:B2)", "3")]
    [InlineData("=COUNT(A1:B2)", "2")]
    [InlineData("=IF(TRUE,1,2)", "1")]
    [InlineData("=IF(FALSE,1,2)", "2")]
    [InlineData("=TRUE()", "TRUE")]
    [InlineData("=FALSE()", "FALSE")]
    [InlineData("=ISNUMBER(\"1\")", "FALSE")]
    [InlineData("=ISNUMBER(A1)", "TRUE")]
    [InlineData("=ISTEXT(B1)", "TRUE")]
    [InlineData("=A1=1", "TRUE")]
    [InlineData("=B1=\"hello\"", "TRUE")]
    [InlineData("=A1<>2", "TRUE")]
    [InlineData("=SUM(Other!A1:A2)", "30")]
    [InlineData("=SUM(MyRange,Other!A1:A2)", "36")]
    [InlineData("=COUNT(MyRange)", "3")]
    [InlineData("=SUM(A1:A3)/COUNT(A1:A3)", "2")]
    [InlineData("=IFERROR(1,2)", "1")]
    [InlineData("=IFERROR(NA(),\"x\")", "x")]
    [InlineData("=ROUND(2.5,0)", "3")]
    [InlineData("=ROUND(-2.5,0)", "-3")]
    [InlineData("=ROUNDUP(1.21,1)", "1.3")]
    [InlineData("=ROUNDDOWN(1.29,1)", "1.2")]
    [InlineData("=MOD(-7,3)", "2")]
    [InlineData("=TEXT(DATE(2024,1,15),\"yyyy/m/d\")", "2024/1/15")]
    [InlineData("=TEXT(0.5,\"0%\")", "50%")]
    [InlineData("=CONCATENATE(\"a\",\"b\")", "ab")]
    [InlineData("=LEFT(\"abc\")", "a")]
    [InlineData("=RIGHT(\"abc\")", "c")]
    [InlineData("=VLOOKUP(2,A1:B3,2,FALSE)", "World")]
    [InlineData("=VLOOKUP(2.5,A1:B3,2,TRUE)", "World")]
    [InlineData("=VLOOKUP(\"World\",B1:C3,2,FALSE)", "6")]
    [InlineData("=VLOOKUP(\"W*\",B1:C3,2,FALSE)", "6")]
    [InlineData("=INDEX(A1:A3,2)", "2")]
    [InlineData("=INDEX(A1:B3,2,2)", "World")]
    [InlineData("=INDEX(A1:B3,0,2)", "Hello")]
    [InlineData("=MATCH(2,A1:A3,0)", "2")]
    [InlineData("=MATCH(2.5,A1:A3,1)", "2")]
    [InlineData("=MATCH(\"World\",B1:B3,0)", "2")]
    [InlineData("=SUMIF(A1:A3,\">1\",K1:K3)", "50")]
    [InlineData("=SUMIF(A1:A3,\"<>2\",K1:K3)", "40")]
    [InlineData("=COUNTIF(A1:A3,\">1\")", "2")]
    [InlineData("=COUNTIF(B1:B3,\"*o*\")", "2")]
    [InlineData("=COUNTIF(A1:A5,\"\")", "2")]
    [InlineData("=INDIRECT(\"A1\")", "1")]
    [InlineData("=INDIRECT(\"Other!A1\")", "10")]
    [InlineData("=INDIRECT(\"'Other'!A1\")", "10")]
    [InlineData("=SUM(INDIRECT(\"A1:A3\"))", "6")]
    [InlineData("=SUM(INDIRECT(\"MyRange\"))", "6")]
    [InlineData("=INDIRECT(\"A\"&\"1\")", "1")]
    [InlineData("=INDIRECT(\"A1:B2\")", "1")]
    [InlineData("=INDIRECT(\"A1\",TRUE)", "1")]
    [InlineData("=ADDRESS(1,1)", "$A$1")]
    [InlineData("=ADDRESS(1,1,4)", "A1")]
    [InlineData("=ADDRESS(2,3,2)", "C$2")]
    [InlineData("=ADDRESS(2,3,3)", "$C2")]
    [InlineData("=ADDRESS(2,3,4)", "C2")]
    [InlineData("=ADDRESS(2,3,1,FALSE)", "R2C3")]
    [InlineData("=ADDRESS(2,3,2,FALSE)", "R2C[3]")]
    [InlineData("=ADDRESS(2,3,3,FALSE)", "R[2]C3")]
    [InlineData("=ADDRESS(2,3,4,FALSE)", "R[2]C[3]")]
    [InlineData("=ADDRESS(1,1,1,TRUE,\"Sheet 1\")", "'Sheet 1'!$A$1")]
    [InlineData("=ADDRESS(1,1,1,TRUE,\"Other\")", "Other!$A$1")]
    [InlineData("=ROW(A5)", "5")]
    [InlineData("=ROW(A5:B7)", "5")]
    [InlineData("=ROW(MyRange)", "1")]
    [InlineData("=COLUMN(C3)", "3")]
    [InlineData("=COLUMN(A5:B7)", "1")]
    [InlineData("=COLUMN(MyRange)", "1")]
    [InlineData("=ROWS(A1:B3)", "3")]
    [InlineData("=COLUMNS(A1:B3)", "2")]
    [InlineData("=ROWS(A1)", "1")]
    [InlineData("=COLUMNS(A1)", "1")]
    [InlineData("=SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", "a+b+c")]
    [InlineData("=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", "a-b+c")]
    [InlineData("=SUBSTITUTE(\"abc\",\"x\",\"y\")", "abc")]
    [InlineData("=SUBSTITUTE(\"hello\",\"l\",\"L\",2)", "helLo")]
    [InlineData("=REPLACE(\"abcdef\",2,3,\"X\")", "aXef")]
    [InlineData("=REPLACE(\"abc\",1,0,\"X\")", "Xabc")]
    [InlineData("=REPLACE(\"abc\",4,1,\"X\")", "abcX")]
    [InlineData("=REPLACE(\"hello\",2,2,\"\")", "hlo")]
    public void Evaluates(string expr, string expected)
    {
        Assert.Equal(expected, Eval("eval.xlsx", "Data", expr));
    }

    [Theory]
    [InlineData("=C1", "3")]
    [InlineData("=C2", "6")]
    [InlineData("=C3", "yes")]
    [InlineData("=D1", "1x")]
    [InlineData("=G2", "2024")]
    [InlineData("=J2", "6")]
    [InlineData("=J3", "12")]
    [InlineData("=J4", "10")]
    [InlineData("=C1+C2", "9")]
    [InlineData("=SUM(C1:C2)", "9")]
    public void FormulaCells(string expr, string expected)
    {
        Assert.Equal(expected, Eval("eval.xlsx", "Data", expr));
    }

    [Theory]
    [InlineData("=A2", "2")]
    [InlineData("=A3", "4")]
    [InlineData("=B1", "7")]
    [InlineData("=B2", "11")]
    [InlineData("=B3", "10")]
    [InlineData("=D1", "11")]
    [InlineData("=D2", "22")]
    [InlineData("=E1", "25")]
    [InlineData("=F1", "35")]
    [InlineData("=SUM(B1:B2)", "18")]
    [InlineData("=Other!A1", "11")]
    public void GraphOrder(string expr, string expected)
    {
        Assert.Equal(expected, Eval("graph.xlsx", "Graph", expr));
    }

    [Theory]
    [InlineData("=C1", "#REF!")]
    [InlineData("=C2", "#REF!")]
    [InlineData("=C3", "#REF!")]
    [InlineData("=SUM(C1:C2)", "#REF!")]
    [InlineData("=Other!A2", "#REF!")]
    public void GraphCycles(string expr, string expected)
    {
        Assert.Equal(expected, Eval("graph.xlsx", "Graph", expr));
    }

    [Theory]
    [InlineData("=A1", "02")]
    [InlineData("=A2", "2")]
    [InlineData("=A3", "2")]
    [InlineData("=A4", "02")]
    [InlineData("=A5", "2")]
    [InlineData("=A6", "2")]
    [InlineData("=A7", "02")]
    [InlineData("=B1", "02")]
    [InlineData("=B2", "2")]
    [InlineData("=B3", "2")]
    [InlineData("=B4", "TRUE")]
    [InlineData("=B5", "TRUE")]
    [InlineData("=TEXT(2,\"00\")", "02")]
    [InlineData("=TEXT(2,\"#,##0.00\")", "2.00")]
    [InlineData("=TEXT(A2,\"00\")", "02")]
    [InlineData("=VALUE(\"02\")", "2")]
    [InlineData("=ISTEXT(\"02\")", "TRUE")]
    [InlineData("=ISNUMBER(02)", "TRUE")]
    [InlineData("=A1+0", "2")]
    public void TextVsNumber(string expr, string expected)
    {
        Assert.Equal(expected, Eval("types.xlsx", "Types", expr));
    }

    [Fact]
    public void NumberFormatTextVsGeneral()
    {
        Assert.Equal("02", NumFmt.FormatText("02", "@"));
        Assert.Equal("02", NumFmt.FormatText("02", "General"));
        Assert.Equal("02", NumFmt.FormatText("02", "00"));
        Assert.Equal("2", NumFmt.Format(2, "General", false));
        Assert.Equal("02", NumFmt.Format(2, "00", false));
    }

    [Theory]
    [InlineData("=NA()", "#N/A")]
    [InlineData("=1/0", "#DIV/0!")]
    [InlineData("=UNKNOWN(1)", "#NAME?")]
    [InlineData("=VLOOKUP(99,A1:B3,2,FALSE)", "#N/A")]
    [InlineData("=VLOOKUP(2,A1:B3,0,FALSE)", "#VALUE!")]
    [InlineData("=VLOOKUP(2,A1:B3,3,FALSE)", "#REF!")]
    [InlineData("=INDEX(A1:A3,4)", "#REF!")]
    [InlineData("=MATCH(\"x\",A1:A3,0)", "#N/A")]
    [InlineData("=MATCH(2,A1:B3,0)", "#N/A")]
    [InlineData("=SUMIF(A1:A3,\">1\",A1:A2)", "#VALUE!")]
    [InlineData("=INDIRECT(\"Other!A1\",FALSE)", "#REF!")]
    [InlineData("=INDIRECT(\"[Book]Sheet!A1\")", "#REF!")]
    [InlineData("=ADDRESS(0,1)", "#VALUE!")]
    [InlineData("=ADDRESS(1,1,5)", "#VALUE!")]
    [InlineData("=SUBSTITUTE(\"aaa\",\"a\",\"b\",0)", "#VALUE!")]
    [InlineData("=REPLACE(\"abc\",0,1,\"X\")", "#VALUE!")]
    [InlineData("=REPLACE(\"abc\",2,-1,\"X\")", "#VALUE!")]
    public void Errors(string expr, string expected)
    {
        Assert.Equal(expected, Eval("eval.xlsx", "Data", expr));
    }

    [Theory]
    [InlineData("=A1+")]
    [InlineData("=SUM(")]
    [InlineData("=1+*2")]
    public void ParseErrors(string expr)
    {
        WithContext("eval.xlsx", "Data", ctx => Assert.Throws<EvalParseException>(() => ctx.Evaluate(expr)));
    }

    [Fact]
    public void OtherSheetContext()
    {
        Assert.Equal("10", Eval("eval.xlsx", "Other", "=A1"));
    }

    [Fact]
    public void XlsValues()
    {
        Assert.Equal("3", Eval("eval.xls", "Data", "=A1+A2"));
        Assert.Equal("6", Eval("eval.xls", "Data", "=SUM(A1:A3)"));
        Assert.Equal("10", Eval("eval.xls", "Data", "=Other!A1"));
        Assert.Equal("Hello", Eval("eval.xls", "Data", "=B1"));
    }

    [Fact]
    public void XlsbNames()
    {
        Assert.Equal("col1", Eval("eval.xlsb", "Named Ranges", "=A1"));
        Assert.Equal("15", Eval("eval.xlsb", "Named Ranges", "=SUM(col1_)"));
    }

    [Fact]
    public void JsonValue()
    {
        WithContext("eval.xlsx", "Data", ctx =>
        {
            EvalValue value = ctx.Evaluate("=A1");
            StringBuilder sb = new StringBuilder();
            EvalJson.AppendValue(sb, value);
            Assert.Equal("1", sb.ToString());
        });
    }
}
