using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Elfo.Console.SAP.XMLtoSSLX
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var app = new Excel.Application();

            string[] files = Directory.GetFiles("D:\\Projects\\Elfo.Console.SAP.CSVtoXML\\bin\\Debug\\net8.0\\OUT\\");

            string directoryOut = "D:\\Projects\\Elfo.Console.SAP.CSVtoXML\\bin\\Debug\\net8.0\\OUT\\XSLX\\";

            //console log delle modifiche

            Directory.CreateDirectory(directoryOut);

            foreach (string file in files)
            {
                var wb1 = app.Workbooks.OpenXML(file);

                string fileNameOutput = directoryOut + Path.GetFileNameWithoutExtension(file) + ".xlsx";

                wb1.SaveAs(fileNameOutput, Excel.XlFileFormat.xlOpenXMLWorkbook);

                wb1.Close();
            }
            app.Quit();
        }
    }
}
