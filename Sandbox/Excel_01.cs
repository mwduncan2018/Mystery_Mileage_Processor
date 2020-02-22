using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace Mystery_1051.Sandbox
{
    public class Excel_01
    {

        public void Run()
        {
            Excel.Application excel = new Excel.Application();



            Excel.Workbook workbook = excel.Workbooks.Open(@"c:\_Mystery_Practice_01\Sandbox\test_01.xlsx", ReadOnly: false, Editable: true);
            Excel.Worksheet worksheet = workbook.Worksheets.Item[1] as Excel.Worksheet;
            if (worksheet == null)
                return;

            var abc = worksheet.Cells[2, 1].Value;
            Excel.Range row1 = worksheet.Rows.Cells[1, 1];
            Excel.Range row2 = worksheet.Rows.Cells[2, 1];

            row1.Value = "Test100";
            row2.Value = "Test200";


            excel.Application.ActiveWorkbook.Save();
            excel.Application.Quit();
            excel.Quit();
        }


    }
}
