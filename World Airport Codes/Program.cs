using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace World_Airport_Codes
{
    class Program
    {
        private static string sep(string s)
        {
            int l = s.IndexOf(" (");
            if (l > 0)
            {
                return s.Substring(0, l);
            }
            return "";

        }

        private static void ExportDataSet(DataSet ds, string destination)
        {
            using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                foreach (System.Data.DataTable table in ds.Tables)
                {

                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                    sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                    DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    uint sheetId = 1;
                    if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                    {
                        sheetId =
                            sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                    List<String> columns = new List<string>();
                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }


                    sheetData.AppendChild(headerRow);

                    foreach (System.Data.DataRow dsrow in table.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        foreach (String col in columns)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }

                }
            }
        }

        private static DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
            TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

        static void Main(string[] args)
        {
            string page = string.Empty;
            string[] letters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            List<string> airports = new List<string>();
            List<Airport> ExportData = new List<Airport>();

            foreach (string letter in letters)
            {
                string addressUrl = String.Format("https://www.world-airport-codes.com/alphabetical/airport-code/{0}.html", letter.ToLower());
                using (WebClient client = new WebClient())
                {
                    page = client.DownloadString(addressUrl);

                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(page);

                    HtmlAgilityPack.HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//tbody/tr/th/a");
                    foreach (var node in nodes)
                    {
                        airports.Add(node.Attributes["href"].Value);
                    }

                }

            }

            double cnt = Convert.ToDouble(airports.Count);
            int i = 0;

            foreach (string airport in airports)
            {
                i++;
                Console.WriteLine("Processing " + airport + "\n" + 
                    i + " out of " + cnt +
                    "\ncurrently at " + String.Format("{0:0.####} %", i / cnt * 100));

                string addressUrl = String.Format("https://www.world-airport-codes.com{0}", airport.ToLower());

                try
                {
                        
                    Airport air = new Airport();

                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();

                    StreamReader reader = new StreamReader(WebRequest.Create(addressUrl).GetResponse().GetResponseStream(), Encoding.UTF8); //put your encoding            
                    doc.Load(reader);


                    air.AirportName = sep(doc.DocumentNode.SelectSingleNode("//h1").InnerText);
                    air.Country = sep(doc.DocumentNode.SelectSingleNode("//h3").InnerText);

                    HtmlAgilityPack.HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//tbody");

                    foreach (HtmlAgilityPack.HtmlNode node in nodes)
                    {

                        foreach (HtmlAgilityPack.HtmlNode row in node.SelectNodes("tr").Descendants("th"))
                        {
                            //Console.WriteLine("row " + row.InnerText);
                            //Console.WriteLine("cell: " + row.NextSibling.InnerText);

                            switch (row.InnerText)
                            {
                                case "IATA Code":
                                    air.IATA = row.NextSibling.InnerText;
                                    break;
                                case "ICAO Code":
                                    air.ICAO = row.NextSibling.InnerText;
                                    break;
                                case "FAA Code":
                                    air.FAA = row.NextSibling.InnerText;
                                    break;
                                case "Latitude":
                                    air.Latitude = row.NextSibling.InnerText;
                                    break;
                                case "Longitude":
                                    air.Longitude = row.NextSibling.InnerText;
                                    break;
                            }

                        }
                    }

                    ExportData.Add(air);
                }
                catch (Exception e)
                {

                    Console.WriteLine(e.Message);
                }

                

            }

            DataTable dt = ConvertToDataTable(ExportData);
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            ExportDataSet(ds, "C:\\world-airport-codes.xlsx");
            
        }
    }
}
