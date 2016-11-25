using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenMcdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VbaReader.Data;

namespace WorkbookModifier
{
    /// <summary>
    /// THIS IS A WORK IN PROGRESS / SCRATCHPAD KIND OF PROJECT,
    /// testing to replace Vba parts in workbooks
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("This tool will replace the existing vbProject of an Excel file with another vbProject");
            Console.WriteLine();

            Console.WriteLine("Enter path of workbook to open: ");
            string path = Console.ReadLine();

            if (!File.Exists(path))
                throw new FileNotFoundException(String.Format("File {0} does not exist", path));

            Console.WriteLine("Enter path of .bin file to open");
            string binPath = Console.ReadLine();

            if (!File.Exists(binPath))
                throw new FileNotFoundException(String.Format("File {0} does not exist", binPath));

            using (SpreadsheetDocument wb = SpreadsheetDocument.Open(path, true))
            {
                var wbPart = wb.WorkbookPart;

                DebugPrintVbaPartOf(wbPart);

                Console.WriteLine();
                Console.WriteLine("---------------REPLACING VBA PART----------------");
                Console.WriteLine("Press any key to continue");
                Console.ReadKey();
                Console.WriteLine();

                // Replace parts
                var vba = wbPart.GetPartsOfType<VbaProjectPart>().FirstOrDefault();
                if(vba != null)
                    wbPart.DeletePart(vba);

                
                VbaProjectPart newVbaPart = wbPart.AddNewPart<VbaProjectPart>();
                using (var stream = File.OpenRead(binPath))
                {
                    newVbaPart.FeedData(stream);
                }

                DebugPrintVbaPartOf(wbPart);

                wbPart.Workbook.Save();

                /*

                foreach(var ws in wbPart.WorksheetParts)
                {
                    Console.WriteLine(ws.Worksheet.LocalName);
                    var sheetData = ws.Worksheet.GetFirstChild<SheetData>();
                }*/
            }

            Console.WriteLine("done");
            Console.ReadKey();
        }

        protected static void DebugPrintVbaPartOf(WorkbookPart wbPart)
        {
            var allParts = wbPart.GetPartsOfType<VbaProjectPart>().ToArray();
            int partsCount = allParts.Count();
            var vba = allParts.FirstOrDefault();

            if (vba == null)
                return;

            var stream = vba.GetStream();
            VbaStorage vbaStorage = new VbaStorage(stream);
            DebugStorage(vbaStorage);
        }

        protected static void DebugStorage(VbaStorage VbaStorage)
        {
            Console.WriteLine("- - - INFO - - -");
            Console.WriteLine("Project name: {0}", VbaStorage.DirStream.InformationRecord.NameRecord.GetProjectNameAsString());
            foreach (KeyValuePair<string, ModuleStream> ms in VbaStorage.ModuleStreams)
            {
                Console.WriteLine("Module stream: {0}", ms.Key);

                Console.WriteLine("Source code:");
                Console.WriteLine(ms.Value.GetUncompressedSourceCodeAsString());
            }


        }

        public static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }
    }
}
