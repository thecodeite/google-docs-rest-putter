using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Spreadsheets;

namespace GoogleData
{
    class Program
    {
        static void Main(string[] args)
        {
            var user = ConfigurationManager.AppSettings["user"];
            var password = ConfigurationManager.AppSettings["password"];
            var spreadSheetName = ConfigurationManager.AppSettings["spreadSheetName"];
            var worksheetName = ConfigurationManager.AppSettings["worksheetName"];
            var dataRange = ConfigurationManager.AppSettings["dataRange"];
            var isisAuth = ConfigurationManager.AppSettings["isisAuth"];
            var consumerId = ConfigurationManager.AppSettings["consumerId"];

            var service = new SpreadsheetsService("codeite-SpreadsheetRestReader-1");
            if (string.IsNullOrWhiteSpace(password))
            {
                Console.WriteLine("password:");
                password = GetPassword();
            }
            service.setUserCredentials(user, password);

            var qqquery = new SpreadsheetQuery();
            Console.WriteLine("Reading spreadsheets");
            SpreadsheetFeed qqfeed = service.Query(qqquery);

            
            var map = new Dictionary<int, SpreadsheetEntry>();
            int index = 1;
            foreach (SpreadsheetEntry ssentry in qqfeed.Entries)
            {
                map[index] = ssentry;
                if (string.IsNullOrWhiteSpace(spreadSheetName)) 
                    Console.WriteLine(index + "] " + ssentry.Title.Text);
                index++;

            }

            SpreadsheetEntry entry;
            if (string.IsNullOrWhiteSpace(spreadSheetName))
            {
                Console.WriteLine("Which spreadsheet:");
                var num = int.Parse(Console.ReadLine()); ;
                entry = map[num];
            }
            else
            {
                var num = map.First(x => x.Value.Title.Text == spreadSheetName).Key;
                entry = map[num];
            }

            AtomLink link = entry.Links.FindService(GDataSpreadsheetsNameTable.WorksheetRel, null);

            WorksheetQuery query = new WorksheetQuery(link.HRef.ToString());
            Console.WriteLine("Reading worksheets");
            WorksheetFeed feed = service.Query(query);

            index = 1;
            var worksheepMap = new Dictionary<int, WorksheetEntry>();
            foreach (WorksheetEntry worksheet in feed.Entries)
            {
                worksheepMap[index] = worksheet;
                if (string.IsNullOrWhiteSpace(worksheetName)) 
                    Console.WriteLine(index + "] " + worksheet.Title.Text);
                index++;
            }

            int workSheenNum;
            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                Console.WriteLine("Which worksheet:");
                workSheenNum = int.Parse(Console.ReadLine());
            }
            else
            {
                workSheenNum = worksheepMap.First(x => x.Value.Title.Text == worksheetName).Key;
            }

            var worksheetentry = worksheepMap[workSheenNum];

            AtomLink cellFeedLink = worksheetentry.Links.FindService(GDataSpreadsheetsNameTable.CellRel, null);


            CellQuery cellQuery = new CellQuery(cellFeedLink.HRef.ToString());
            ReadDataRange(dataRange, cellQuery);

            Console.WriteLine("Reading cells");
            CellFeed cellfeed = service.Query(cellQuery);

            var cells = new Dictionary<uint, Dictionary<uint, string>>();
            foreach (CellEntry curCell in cellfeed.Entries)
            {
                cells[curCell.Cell.Row] = cells.GetOrNull(curCell.Cell.Row) ?? new Dictionary<uint, string>();

                cells[curCell.Cell.Row][curCell.Cell.Column] = curCell.Value;

            }

            foreach (var cell in cells.Values)
            {
                var url = cell[7];
                var uri = new Uri(url);
                var data = cell[8];
                Console.WriteLine("Posting {1} bytes to {0}", uri, data.Length);


                var request = new HttpRequestFactory().Create(uri);
                request.Method = "PUT";
                request.Headers["Authorization"] = isisAuth;
                request.Headers["Consumer-Id"] = consumerId;

                using (var requestStream = request.GetRequestStream())
                using (var writer = new StreamWriter(requestStream))
                {
                    writer.Write(data);
                }

                HttpWebResponse response;
                try
                {
                    response = request.GetResponse() as HttpWebResponse;
                }
                catch (WebException ex)
                {
                    if (ex.Response == null || ex.Status != WebExceptionStatus.ProtocolError)
                        throw;

                    response = (HttpWebResponse)ex.Response;
                }

                Console.WriteLine(response.StatusCode);
                using (var responseStream = response.GetResponseStream())
                {
                    if (responseStream != null)
                        Console.WriteLine(new StreamReader(responseStream).ReadToEnd());
                }

            }

            Console.WriteLine("Done");
            Console.ReadKey();
        }

        private static void ReadDataRange(string dataRange, CellQuery cellQuery)
        {
            var regex = new Regex(@"(\d+)-(\d+),(\d+)-(\d+)");

            var match = regex.Match(dataRange);

            if (!match.Success)
            {
                throw new Exception("Data range should be in format <rowStart>-<rowEnd>,<colStart>-<colEnd>");
            }

            cellQuery.MinimumRow = uint.Parse(match.Groups[1].Value);
            cellQuery.MaximumRow = uint.Parse(match.Groups[2].Value);
            cellQuery.MinimumColumn = uint.Parse(match.Groups[3].Value);
            cellQuery.MaximumColumn = uint.Parse(match.Groups[4].Value);

        }

        private static string GetPassword()
        {
            var pwd = new StringBuilder();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.Remove(pwd.Length - 1, 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    pwd.Append(i.KeyChar);
                    Console.Write("*");
                }
            }

            Console.WriteLine();
            return pwd.ToString();
        }
    }

    public static class DictionaryHelper
    {
        public static TV GetOrNull<TK, TV>(this Dictionary<TK, TV> dictionary, TK key) where TV : class
        {
            TV value;
            return dictionary.TryGetValue(key, out value) ? value : null;
        }
    }
}
