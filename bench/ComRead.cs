using System;
using System.Diagnostics;
using System.Reflection;

// Minimal late-bound Excel COM reader used only for benchmarking.
// Requires Microsoft Excel (Excel.Application) to be installed.
class ComRead
{
    static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine("usage: ComRead.exe <workbook>");
            Environment.Exit(2);
        }

        Type excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
        {
            Console.Error.WriteLine("Excel.Application is not registered");
            Environment.Exit(2);
        }

        string path = System.IO.Path.GetFullPath(args[0]);
        Stopwatch sw = Stopwatch.StartNew();
        object app = null;
        try
        {
            app = Activator.CreateInstance(excelType);
            excelType.InvokeMember("Visible", BindingFlags.SetProperty, null, app, new object[] { false });
            excelType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, app, new object[] { false });

            object workbooks = excelType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, app, null);
            Type workbooksType = workbooks.GetType();
            object book = workbooksType.InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks,
                new object[] { path });

            object sheet = excelType.InvokeMember("ActiveSheet", BindingFlags.GetProperty, null, app, null);
            object used = sheet.GetType().InvokeMember("UsedRange", BindingFlags.GetProperty, null, sheet, null);
            object value = used.GetType().InvokeMember("Value2", BindingFlags.GetProperty, null, used, null);

            long count = 0;
            Array arr = value as Array;
            if (arr != null)
            {
                foreach (object o in arr)
                {
                    if (o != null && !(o is DBNull)) count++;
                }
            }

            sw.Stop();
            Console.WriteLine("{\"count\":" + count + ",\"ms\":" + sw.ElapsedMilliseconds + "}");

            book.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, book, new object[] { false });
        }
        finally
        {
            if (app != null)
            {
                try { excelType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, app, null); }
                catch { }
            }
        }
    }
}
