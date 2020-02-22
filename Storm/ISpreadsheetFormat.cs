using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mystery_1051.Storm
{
    public interface ISpreadsheetFormat
    {
        bool HasRows();
        string GetCurrentAddress();
        string GetNextAddress();
        void RecordMileage(float mileage);
        void NextRow();

        void StartUp();
        void ShutDown();

    }
}
