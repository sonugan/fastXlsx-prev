using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace excelbig
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Temp\";
            var fileName = Path.Combine(path, @"test.xlsx");
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }


            var headerList = new string[] { "Header 1", "Header 2", "Header 3", "Header 4" };
            //sheet1 
            var boolList = new bool[] { true, false, true, false };
            var intList = new int[] { 1, 2, 3, -4 };
            var dateList = new DateTime[] { DateTime.Now, DateTime.Today, DateTime.Parse("1/1/2014"), DateTime.Parse("2/2/2014") };
            var sharedStringList = new string[] { "shared string", "shared string", "cell 3", "cell 4" };
            var inlineStringList = new string[] { "inline string", "inline string", "3>", "<4" };
            
            using (var spreadSheet = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // create the workbook
                var workbookPart = spreadSheet.AddWorkbookPart();

                var openXmlExportHelper = new OpenXmlWriterHelper();
                //openXmlExportHelper.SaveCustomStylesheet(workbookPart);
                var style = new XlsxStyleSheet(workbookPart);

                style.AddCellFormat(new XlsxCellFormat() { Name = "Data", ForegroundColor = new XlsxColor("bb4055") });
                style.AddCellFormat(new XlsxCellFormat()
                {
                    Name = "Header",
                    ForegroundColor = new XlsxColor("C8EEFF"),
                    Border = XlsxBorder.CreateBox(BorderStyleValues.Medium, new XlsxColor("bb4055"))
                });
                style.AddCellFormat(new XlsxCellFormat() { Name = "DataDate", ForegroundColor = new XlsxColor("bb4055"), NumberFormat = @"[$-409]m/d/yy\ h:mm\ AM/PM;@" });

                style.Save();


                var workbook = workbookPart.Workbook = new Workbook();
                var sheets = workbook.AppendChild<Sheets>(new Sheets());

                
                // create worksheet 1
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                sheets.Append(sheet);
                
                using (var worksheet = new XlsxWorksSheet(style, workbookPart, worksheetPart))
                {
                    worksheet.WriteRow(new List<XlsxCell>()
                        {
                            new XlsxSharedStringCell("Id", "Header"),
                            new XlsxSharedStringCell("Name", "Header"),
                            new XlsxSharedStringCell("Lastname", "Header"),
                            new XlsxSharedStringCell("DocumentNumber", "Header"),
                            new XlsxSharedStringCell("Birthdate", "Header")
                        });

                    var odd = true;
                    foreach(var data in GenerateData())
                    {
                        var format = "";
                        var formatDate = "";
                        if (!odd)
                        {
                            format = "Data";
                            formatDate = "DataDate";
                            odd = true;
                        }
                        else
                        {
                            odd = false;
                        }
                        worksheet.WriteRow(new List<XlsxCell>()
                        {
                            new XlsxSharedStringCell(data.Id.ToString(), format),
                            new XlsxSharedStringCell(data.Name, format),
                            new XlsxSharedStringCell(data.Lastname, format),
                            new XlsxSharedStringCell(data.DocumentNumber, format),
                            new XlsxDateCell(data.Birthdate.ToOADate().ToString(CultureInfo.InvariantCulture), formatDate)
                        });
                    }
                }
            }
        }

        public static List<Data> GenerateData()
        {
            var data = new List<Data>();
            var random = new Random((int)DateTime.Now.Ticks);

            for(var i = 0; i < 100; i++)
            {
                data.Add(new Data()
                {
                    Id = i,
                    Name = "asdfasdfasdfasdfasdfasdfs",
                    Lastname = "asdfasdfasdfasdfasdfasdfasdf",
                    DocumentNumber = "asdfasdfasdfasdfasdf",
                    Birthdate = DateTime.Now
                });
            }

            return data;
        }
    }

    public class Data
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Lastname { get; set; }
        public string DocumentNumber { get; set; }
        public DateTime Birthdate { get; set; }
    }
}