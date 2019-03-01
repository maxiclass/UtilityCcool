using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using System.IO;


namespace FinaleAppUtility
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}

namespace PanelFunctions
{
    static class ExcelPanelFunctions
    {
        public static void OpenExcel()
        {
         Excel.Application excel = new Excel.Application();
         Excel.Workbook wb = excel.Workbooks.Open(@"C:\\Users\\uia93155\\Desktop\\timeprj\\FinaleAppUtility\\Database.xlsx");
         Excel.Worksheet sheet = (Excel.Worksheet)wb.Sheets["DataTable"];

            ((Excel.Range)sheet.Cells[1,2]).Value = "Hellox";
            wb.Save();
           //wb.SaveAs(@"C:\\Users\\uia93155\\Desktop\\timeprj\\FinaleAppUtility\\Database.xlsx");
            wb.Close();
            excel.Quit();
        }

        public static string ReadExcel(int row, int col)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(@"C:\\Users\\uia93155\\Desktop\\timeprj\\FinaleAppUtility\\Database.xlsx");
            Excel.Worksheet sheet = (Excel.Worksheet)wb.Sheets["DataTable"]; 
            var cell = ((Excel.Range)sheet.Cells[row, col]);
            string strExcelRead;
            strExcelRead = cell.Text;
            wb.Save();
            //wb.SaveAs(@"C:\\Users\\uia93155\\Desktop\\timeprj\\FinaleAppUtility\\Database.xlsx"); DataTable
            wb.Close();
            excel.Quit();
            return strExcelRead;
        }

        public static void WriteExcelCell(int row, int col)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(@"C:\\Users\\uia93155\\Desktop\\timeprj\\FinaleAppUtility\\Database.xlsx");
            Excel.Worksheet sheet = (Excel.Worksheet)wb.Sheets["InfoData"];
            ((Excel.Range)sheet.Cells[row, col]).Value = "Hellox";

            wb.Save();
            //wb.SaveAs(@"C:\\Users\\uia93155\\Desktop\\timeprj\\FinaleAppUtility\\Database.xlsx");
            wb.Close();
            excel.Quit();
        }

    }
}