using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelInteropTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var excel = new Excel.Application();
            excel.Visible = true;
            var wb = excel.Workbooks.Add();

            VBProject vba = wb.VBProject;
            var module = vba.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);

            module.CodeModule.AddFromString(
                "Public Sub Test() " + Environment.NewLine +
                "   MsgBox \"Hello\"" + Environment.NewLine +
                "End Sub");

            Console.WriteLine("waiting for input...");
            Console.ReadKey();

            Marshal.ReleaseComObject(module);
            module = null;
            Marshal.FinalReleaseComObject(vba);
            vba = null;
            Marshal.ReleaseComObject(wb);
            wb = null;
            Marshal.ReleaseComObject(excel);
            excel = null;

            excel.Quit();
        }
    }
}
