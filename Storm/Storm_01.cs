using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Mystery_1051.Storm
{
    public class ExcelMileageUpdater
    {
        //Method used to get the mileage
        public MileageMethod MileageMethod;
        //Where to read the file on the hard drive
        public string ExcelFileReadLocation;
        //Where to write the modified file on the hard drive
        public string ExcelFileWriteLocation;

        //Constructor that uses Dependency Injection
        public ExcelMileageUpdater(MileageMethod mileageMethod, 
            string excelFileReadLocation, 
            string excelFileWriteLocation)
        {
            MileageMethod = mileageMethod;
            ExcelFileReadLocation= excelFileReadLocation;
            ExcelFileWriteLocation = excelFileWriteLocation;
        }

        public void CreateNewExcelFile()
        {
            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
                Console.WriteLine("uh oh");
            else
                Console.WriteLine("f!");

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "shit";
            xlWorkSheet.Cells[1, 2] = "Name";
            xlWorkSheet.Cells[2, 1] = "1";
            xlWorkSheet.Cells[2, 2] = "One";
            xlWorkSheet.Cells[3, 1] = "2";
            xlWorkSheet.Cells[3, 2] = "Two";

            string saveLocation = "c:\\_Mystery_Practice_01\\saved_25.xlsx";

            xlWorkBook.SaveAs(saveLocation,
                Excel.XlFileFormat.xlWorkbookNormal,
                misValue, misValue, misValue, misValue,
                Excel.XlSaveAsAccessMode.xlExclusive,
                misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }


        public void ReadFromExcelFile()
        {
            //read from excel file
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt, cCnt, rw, cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(
                @"c:\_Mystery_Practice_01\p_01.xlsx", 0, true, 5, "", "", true,
                Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    Console.WriteLine(str);
                }
            }

            xlWorkSheet.Cells[1, 1] = "shit";
            string saveLocation = "c:\\_Mystery_Practice_01\\saved_25.xlsx";

            object misValue = System.Reflection.Missing.Value;
            xlWorkBook.SaveAs(saveLocation,
                Excel.XlFileFormat.xlWorkbookNormal,
                misValue, misValue, misValue, misValue,
                Excel.XlSaveAsAccessMode.xlExclusive,
                misValue, misValue, misValue, misValue, misValue);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

//            xlWorkBook.Close(true, null, null);
//            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        public void Run()
        {
            //calculate mileage


            // if B is "H"
            //  read address (I)
            //  if address is not empty
            //      read zipcode (H)
            //      result = send address to GoogleMapsAPI
            //      


            // "W" or "H" signals the start of a new day
            // "W" starts from work address
            // "H" starts from home address
            // No street address indicates a phone shop and to ignore the line
            // "Mileage To" is where to put the mileage from the previous address to the current address
            // "Mileage From" is only used at the very end of the day (on the last line for that day)
            // Mileage From is the distance to drive home from the last shop
            // if "W", then substract "Z" from "Mileage From" and overwrite Mileage From

            //write to excel file
        }

    }

    public interface MileageMethod
    {

    }

    public class GoogleMapsAPI : MileageMethod
    {

    }

    public class SeleniumWebDriver : MileageMethod
    {

    }

}
