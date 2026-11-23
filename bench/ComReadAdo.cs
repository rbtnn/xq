using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

// COM benchmark via ADO + Microsoft ACE OLEDB. This is a COM-based reader
// (ADODB.Connection) and does not require Excel to be activated. It reads
// .xlsx / .xls through the ACE provider installed with Office/Access.
class ComReadAdo
{
    static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine("usage: ComReadAdo.exe <workbook> [sheet]");
            Environment.Exit(2);
        }

        string path = Path.GetFullPath(args[0]);
        string sheet = args.Length > 1 ? args[1] : "Data";
        string ext = Path.GetExtension(path).ToLowerInvariant();
        string props = ext == ".xls" ? "Excel 8.0;HDR=NO" : "Excel 12.0 Xml;HDR=NO";

        Type connType = Type.GetTypeFromProgID("ADODB.Connection");
        if (connType == null)
        {
            Console.Error.WriteLine("ADODB.Connection is not registered");
            Environment.Exit(2);
        }

        Stopwatch sw = Stopwatch.StartNew();
        object conn = null;
        object rs = null;
        string provider = "Microsoft.ACE.OLEDB.16.0";
        try
        {
            conn = Activator.CreateInstance(connType);
            string connStr = "Provider=" + provider + ";Data Source=" + path +
                ";Extended Properties=\"" + props + "\";";
            try
            {
                connType.InvokeMember("Open", BindingFlags.InvokeMethod, null, conn, new object[] { connStr });
            }
            catch
            {
                provider = "Microsoft.ACE.OLEDB.12.0";
                connStr = "Provider=" + provider + ";Data Source=" + path +
                    ";Extended Properties=\"" + props + "\";";
                connType.InvokeMember("Open", BindingFlags.InvokeMethod, null, conn, new object[] { connStr });
            }

            string query = "SELECT * FROM [" + sheet + "$]";
            rs = connType.InvokeMember("Execute", BindingFlags.InvokeMethod, null, conn, new object[] { query });
            Type rsType = rs.GetType();

            long count = 0;
            while (true)
            {
                bool eof = (bool)rsType.InvokeMember("EOF", BindingFlags.GetProperty, null, rs, null);
                if (eof) break;

                object rows = rsType.InvokeMember("GetRows", BindingFlags.InvokeMethod, null, rs, new object[] { 1000 });
                Array arr = rows as Array;
                if (arr == null) break;

                int nrows = arr.GetLength(1);
                int ncols = arr.GetLength(0);
                for (int r = 0; r < nrows; r++)
                {
                    for (int c = 0; c < ncols; c++)
                    {
                        object v = arr.GetValue(c, r);
                        if (v != null && !(v is DBNull)) count++;
                    }
                }
                if (nrows < 1000) break;
            }

            sw.Stop();
            Console.WriteLine("{\"count\":" + count + ",\"ms\":" + sw.ElapsedMilliseconds +
                ",\"engine\":\"ADO/ACE\",\"provider\":\"" + provider + "\"}");

            rsType.InvokeMember("Close", BindingFlags.InvokeMethod, null, rs, null);
        }
        finally
        {
            if (rs != null) { try { rs.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, rs, null); } catch { } }
            if (conn != null) { try { connType.InvokeMember("Close", BindingFlags.InvokeMethod, null, conn, null); } catch { } }
        }
    }
}
