using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mystery_1051.Storm
{
    public class SpreadsheetFormat2019 : ISpreadsheetFormat
    {
        private const string WORK_ADDRESS = "7500 Security Blvd, Woodlawn, Maryland 21244";
        private const string HOME_ADDRESS = "3427 Ady Rd, Street, Maryland 21154";

        private int startRow;
        private int currentRow;
        private int lastRow;

        private SimpleExcel excel;
        public void StartUp() { excel.StartUp(); }
        public void ShutDown() { excel.ShutDown(); }

        public SpreadsheetFormat2019(SimpleExcel excelFile, int start = 4, int last = 597)
        {
            excel = excelFile;
            startRow = start; // first row to process
            lastRow = last; // last row to process

            currentRow = startRow;
        }

        public bool HasRows()
        {
            return (currentRow <= lastRow) ? true : false;
        }

        public string GetCurrentAddress()
        {
            if (excel.ReadCell("A", currentRow) == "P")
            {
                currentRow++;
                return GetCurrentAddress();
            }

            if (excel.ReadCell("A", currentRow) == "W")
            {
                return WORK_ADDRESS;
            }

            if (excel.ReadCell("A", currentRow) == "H")
            {
                return HOME_ADDRESS;
            }

            if (excel.ReadCell("A", currentRow) == "")
            {
                string streetAddress = excel.ReadCell("H", (currentRow - 1));
                string city = excel.ReadCell("F", (currentRow - 1));
                string zip = excel.ReadCell("G", (currentRow - 1));

                string address = streetAddress + ", " + city + " " + zip;
                return address;
            }
            else
            {
                return "ERROR";
            }
        }

        public string GetNextAddress()
        {
            if (excel.ReadCell("E", currentRow) == "")
            {
                return HOME_ADDRESS;
            }

            else
            {
                string streetAddress = excel.ReadCell("H", (currentRow));
                string city = excel.ReadCell("F", (currentRow));
                string zip = excel.ReadCell("G", (currentRow));

                string address = streetAddress + ", " + city + " " + zip;
                return address;
            }
        }

        public void RecordMileage(float mileage)
        {
            if (excel.ReadCell("E", currentRow) != "")
            {
                excel.WriteCell("N", currentRow, mileage.ToString());
            }

            if (excel.ReadCell("E", currentRow) == "")
            {
                excel.WriteCell("O", (currentRow - 1), mileage.ToString());
            }
        }



        public void NextRow()
        {
            currentRow++;
        }

    }
}