using Mystery_1051.Storm;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Mystery_1051
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting Mystery Mileage Processor");

            //string fileRead = @"c:\_Mystery_File_2018\2018-SS-MileageTestFile-7.xlsx";
            string fileRead = @"c:\_Mystery_File_2018\2019-SS-Mileage-Calculation-File.xlsx";
            string fileWrite = @"c:\_Mystery_File_2018\TestResult_";
     
            //arrange
            var readLocation = fileRead;
            var writeLocation = fileWrite + "MysteryTest_" + DateTime.Now.ToString("_H_mm_ss_") + ".xlsx";
            var _sut = new MysteryMileageProcessor(
                new SpreadsheetFormat2019(
                    new SimpleExcel(readLocation, writeLocation)));

            //act
            _sut.Run();

            Console.WriteLine("Complete");
            Console.ReadLine();

        }
    }
}
