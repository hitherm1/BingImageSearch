// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Newtonsoft.Json;
using System;
using System.Text;
using System.Net;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using Newtonsoft.Json.Linq;
using System.Text.Json.Serialization;
using System.Net.Http;
using System.Text.Json;


// This sample uses the Bing Web Search API v7 to retrieve different kinds of media from the web.

namespace BingWebSearch
{
    class Program
    {


        public class WebPage
        {
            [JsonPropertyName("url")]
            public string ThumbnailUrl { get; set; }
        }

        public class WebPages
        {
            [JsonPropertyName("value")]
            public WebPage[] Value { get; set; }
        }

        public class SearchResult
        {
            [JsonPropertyName("webPages")]
            public WebPages WebPages { get; set; }
        }


        // Add your Bing Search V7 subscription key and endpoint to your environment variables
        static string subscriptionKey = "bc2c80741a9940249335a679440846f8";
        static string endpoint = "https://api.bing.microsoft.com/v7.0/search";
        private static string thumbnailUrl;
        private static int i;
        private static string url;
        private static int dim;
        private static string saveid;
        private static object sett;
        private static int result;
        private static object quantity;
        private static object variety;
        private static object extvalue2;

        static void Main()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string excelFilePath = @"C:\Users\10045006\Documents\OpenNumismat\output.xlsx"; // Replace with the path to your Excel file
            string filePath = @"C:\Users\10045006\Documents\OpenNumismat\output2.xlsx"; // Change this to your desired file path




            try
            {
                FileInfo excelFile = new FileInfo(excelFilePath);

                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming the data is in the first worksheet

                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++) // Assuming the data starts from the second row
                    {


                        string data = worksheet.Cells[row, 12].Value?.ToString() + " " + worksheet.Cells[row, 6].Value?.ToString() + " " + worksheet.Cells[row, 5].Value?.ToString() + " " + worksheet.Cells[row, 4].Value?.ToString(); // Retrieving data from the second column

                        Console.WriteLine(worksheet.Cells[row, 2].Value?.ToString());
                        if (!string.IsNullOrEmpty(data))
                            ProcessData(data, worksheet.Cells[row, 1].Value?.ToString(), worksheet.Cells[row, 2].Value?.ToString(), worksheet.Cells[row, 3].Value?.ToString(), worksheet.Cells[row, 4].Value?.ToString(), worksheet.Cells[row, 5].Value?.ToString(), worksheet.Cells[row, 6].Value?.ToString(), worksheet.Cells[row, 7].Value?.ToString(), worksheet.Cells[row, 8].Value?.ToString(), worksheet.Cells[row, 9].Value?.ToString(), worksheet.Cells[row, 10].Value?.ToString(), worksheet.Cells[row, 11].Value?.ToString(), worksheet.Cells[row, 12].Value?.ToString(), worksheet.Cells[row, 13].Value?.ToString(), worksheet.Cells[row, 14].Value?.ToString());
                    }
                }
            }


            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        static void ProcessData(string data, string id, string title, string quantity, string value2, string extvalue2, string unit, string country, string year, string variety, string mintmark, string series, string sett, string status, string url)
        {
            // Implement your method logic here
            // Console.WriteLine($"Processing data: {data}");
            // Create a dictionary to store relevant headers
            Dictionary<String, String> relevantHeaders = new Dictionary<String, String>();

            Console.OutputEncoding = Encoding.UTF8;

            // Console.WriteLine("Searching the Web for: " + query);

            // Construct the URI of the search request

            var uriQuery = endpoint + "?q=" + Uri.EscapeDataString(data);

            // Perform the Web request and get the response
            WebRequest request = HttpWebRequest.Create(uriQuery);
            //Console.WriteLine("data: " +    data);
            request.Headers["Ocp-Apim-Subscription-Key"] = subscriptionKey;
            HttpWebResponse response = (HttpWebResponse)request.GetResponseAsync().Result;
            string json = new StreamReader(response.GetResponseStream()).ReadToEnd();

            JsonDocument doc = JsonDocument.Parse(json);

            JsonElement root = doc.RootElement;
            JsonElement webPages = root.GetProperty("webPages");
            JsonElement valueArray = webPages.GetProperty("value");
            url = string.Empty;
            foreach (JsonElement value in valueArray.EnumerateArray())
            {
                url += value.GetProperty("url").GetString() + "|";
                //Console.WriteLine(url);



                // Console.WriteLine("json: " + json);
                // Extract Bing HTTP headers
                foreach (String header in response.Headers)
                {
                    if (header.StartsWith("BingAPIs-") || header.StartsWith("X-MSEdge-"))
                        relevantHeaders[header] = response.Headers[header];
                }
            }




            // Create a new Excel package and worksheet if the file doesn't exist
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                string filePath = @"C:\Users\10045006\Documents\OpenNumismat\output2.xlsx"; // Change this to your desired file path


                if (i == 0)
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    // Add headers
                    worksheet.Cells[1, 1].Value = "id";
                    worksheet.Cells[1, 2].Value = "title";
                    worksheet.Cells[1, 3].Value = "quantity";
                    worksheet.Cells[1, 4].Value = "value";
                    worksheet.Cells[1, 5].Value = "extended value";
                    worksheet.Cells[1, 6].Value = "unitr";
                    //worksheet.Cells[i, 7].Value = data;
                    worksheet.Cells[1, 7].Value = "country";
                    worksheet.Cells[1, 8].Value = "year";
                    worksheet.Cells[1, 9].Value = "variety";
                    worksheet.Cells[1, 10].Value = "mintmark";
                    worksheet.Cells[1, 11].Value = "series";
                    worksheet.Cells[1, 12].Value = "set";
                    worksheet.Cells[1, 13].Value = "status";
                    worksheet.Cells[1, 14].Value = "url";



                    try
                    {
                        package.SaveAs(new FileInfo(filePath));

                    }
                    catch (Exception ex)
                    {

                    }
                    i = 2;
                }


                // Append data to an existing Excel file


                // Find the last used row
                // int lastUsedRow = worksheet.Dimension.End.Row;
                sett = " ";
                using (var package2 = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet2 = package2.Workbook.Worksheets[0];

                    // Find the last used row
                    int lastUsedRow = worksheet2.Dimension.End.Row;

                    if (series.Contains("USA"))
                    {
                        switch (unit)
                        {
                            case "1 Cent":


                                // Try to parse the string to an integer
                                if (int.TryParse(year, out result))
                                {
                                    if (result >= 1941 && result <= 1974)
                                    {
                                        series = "Coin Sets - USA";
                                        sett = "H E Harris Lincoln Cent 1941-1974";
                                    }
                                    if (result >= 1909 && result <= 1940)
                                    {
                                        series = "Coin Sets - USA";
                                        sett = "H E Harris Lincoln Cent 1909-1940";
                                    }
                                    if (result >= 1975 && result <= 2013)
                                    {
                                        series = "Coin Sets - USA";
                                        sett = "H E Harris Lincoln Cent 1975-2013";
                                    }
                                    if (result >= 1857 && result <= 1909)
                                    {
                                        series = "Coin Sets - USA";
                                        sett = "H E Harris Flying Eagle and Indian Head 1857-1909";
                                    }
                                }
                                break;
                            case "25 Cents":
                                // Try to parse the string to an integer
                                if (int.TryParse(year, out result))
                                {
                                    if (result >= 1999 && result <= 2003)
                                    {
                                        series = "Coin Sets - USA";
                                        sett = "H E Harris Washington Quarters State Collection 1999-2003";
                                    }
                                    if (result >= 2010 && result <= 2021)
                                    {
                                        series = "Coin Sets - USA";
                                        sett = "H E Harris National Park Quarters State Collection 2010-2021";
                                    }

                                }
                                break;
                            case "5 Cents":
                                // Try to parse the string to an integer
                                if (int.TryParse(year, out result))
                                {
                                    if (result >= 1913 && result <= 1938)
                                    {
                                        series = "Coin Sets - USA";
                                        sett = "H E Harris Buffalo Nickel 1913-1938";
                                    }
                                    if (result >= 2010 && result <= 2021)
                                    {
                                        series = "Coin Sets - USA";
                                        sett = "H E Harris National Park Quarters State Collection 2010-2021";
                                    }

                                }
                                break;
                            default:
                                Console.WriteLine("The number is not 1, 2, or 3");
                                break;
                        }

                    }

                    worksheet2.Cells[i, 1].Value = id;
                    worksheet2.Cells[i, 2].Value = title;
                    worksheet2.Cells[i, 3].Value = quantity;
                    worksheet2.Cells[i, 4].Value = value2;
                    worksheet2.Cells[i, 5].Value = extvalue2;

                    worksheet2.Cells[i, 6].Value = unit;
                    worksheet2.Cells[i, 7].Value = country;
                    worksheet2.Cells[i, 8].Value = year;
                    worksheet.Cells[i, 9].Value = variety;

                    if (mintmark != null)
                        worksheet2.Cells[i, 10].Value = mintmark;

                    worksheet2.Cells[i, 11].Value = series;
                    worksheet2.Cells[i, 12].Value = sett;
                    worksheet2.Cells[i, 13].Value = status;
                    worksheet2.Cells[i, 14].Value = url;

                    // Save the package with the appended data
                    try
                    {
                        package2.Save();
                    }
                    catch (Exception ex)
                    {

                    }
                }
                i++;


                //Console.WriteLine("Data has been written to the Excel file.");
            }
        }
    }
}

