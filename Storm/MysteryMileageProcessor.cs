using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

namespace Mystery_1051.Storm
{
    public class MysteryMileageProcessor
    {
        //free Google Maps API Driving Directions
        //2500 max directions per day
        //100 max per second
        //25 max per query
        //old_GOOGLE_MAPS_API_KEY = "AIzaSyCtymGbsYhtmAYJf6u10iilg6dqHfuV2FE";
        private string GOOGLE_MAPS_API_KEY = "AIzaSyCG3enzFWG3-JUtAUxJIOxKDDFrxfJVmFA";
        private ISpreadsheetFormat spreadsheet;

        private int count = 0;

        //Dependency Injection
        //allows easier modification of code if the layout/format of the Excel spreadsheet changes in the future
        public MysteryMileageProcessor(ISpreadsheetFormat spreadsheetFormat)
        {
            spreadsheet = spreadsheetFormat;
        }

        public void Run()
        {
            spreadsheet.StartUp();

            while (spreadsheet.HasRows())
            {
                //Get the current and next addresses
                var currentAddress = spreadsheet.GetCurrentAddress();
                var nextAddress = spreadsheet.GetNextAddress();
                Console.WriteLine(currentAddress + " --- " + nextAddress);

                if (currentAddress != "ERROR")
                {
                    //Create Google Maps Url (Builder Design Pattern)
                    var googleUrl = new GoogleMapsUrlBuilder()
                        .WithOrigin(currentAddress)
                        .WithDestination(nextAddress)
                        .WithKey(GOOGLE_MAPS_API_KEY)
                        .Build();

                    //Get JSON from Google Maps API
                    //...
                    //          var webRequest = WebRequest.Create(googleUrl);
                    //          var httpWebResponse = (HttpWebResponse)webRequest.GetResponse();
                    //          var stream = httpWebResponse.GetResponseStream();
                    //          var streamReader = new StreamReader(stream);
                    //          var jsonString = streamReader.ReadToEnd();
                    //...
                    //OR IN ONE LINE!!
                    var jsonString = (new StreamReader(((HttpWebResponse)WebRequest.Create(googleUrl).GetResponse()).GetResponseStream())).ReadToEnd();
                    count++;
                    System.IO.File.WriteAllText(@"c:\_Mystery_File_2018\jsons\json_" + count.ToString() + ".txt", jsonString);

                    //Deserialize Google Maps JSON into an Object
                    var googleMapsJson = JsonConvert
                        .DeserializeObject<GoogleMapsJson>(jsonString);

                    //Get mileage (string) from the JSON Object
                    string mileageString = "999999999";
                    try
                    {
                        mileageString = googleMapsJson
                            .routes
                            .ToList()[0]
                            .legs
                            .ToList()[0]
                            .distance
                            .text;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }


                    //Convert the mileage (string) into a float using an Extension Method
                    float mileage = mileageString.ConvertMileageStringToFloat();
                    Console.WriteLine(mileage);

                    //Record the mileage
                    spreadsheet.RecordMileage(mileage);
                    Console.WriteLine("Recorded...");
                    Console.WriteLine();

                }
                spreadsheet.NextRow();
            }

            spreadsheet.ShutDown();

        }

    }
}
