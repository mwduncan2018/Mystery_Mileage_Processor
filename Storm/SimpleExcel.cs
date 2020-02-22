using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.CSharp;

namespace Mystery_1051.Storm
{
    public class SimpleExcel
    {
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkBook;
        private Excel.Worksheet xlWorkSheet;
        private Excel.Range xlRange;

        private string fileReadLocation;
        private string fileWriteLocation;


        public SimpleExcel(string fileRead, string fileWrite)
        {
            fileReadLocation = fileRead;
            fileWriteLocation = fileWrite;
        }

        //opens the excel connection
        public void StartUp()
        {
            xlApp = new Excel.Application();

            xlWorkBook = xlApp.Workbooks.Open(
                fileReadLocation);

            xlWorkSheet = xlWorkBook.Sheets[1];
            xlRange = xlWorkSheet.UsedRange;
        }

        //updates the cell with the value
        public void WriteCell(string col, int row, string value)
        {
            Excel.Range cell = xlWorkSheet.Rows.Cells[row, col];

            cell.Value = value;
        }

        public string ReadCell(string col, int row)
        {
            try
            {
                return xlWorkSheet.Cells[row, col].Value.ToString();
            }
            catch (RuntimeBinderException e)
            {
                return "";
            }
        }

        //closes the excel connection
        public void ShutDown()
        {
            xlWorkBook.SaveAs(fileWriteLocation);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(xlRange);
            Marshal.FinalReleaseComObject(xlWorkSheet);

            xlWorkBook.Close(Type.Missing, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(xlWorkBook);

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);

        }

    }
}
