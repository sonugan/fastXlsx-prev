using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelbig
{
    public class XlsxWorksSheet : IDisposable
    {
        public readonly OpenXmlWriter Writer;
        private readonly Dictionary<string, int> ShareStringDictionary = new Dictionary<string, int>();
        private readonly WorkbookPart WorkbookPart;
        private readonly XlsxStyleSheet Style;

        public XlsxWorksSheet(XlsxStyleSheet style, WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            Style = style;
            WorkbookPart = workbookPart;
            Writer = OpenXmlWriter.Create(worksheetPart);
            Writer.WriteStartElement(new Worksheet());
            Writer.WriteStartElement(new SheetData()); // TODO: poner en una clase sheet
        }

        public void WriteRow(IList<XlsxCell> cells)
        {
            Writer.WriteStartElement(new Row());
            for (int i = 0; i < cells.Count(); i++)
            {
                var style = !string.IsNullOrEmpty(cells[i].FormatName) ? Style.CellFormats[cells[i].FormatName] : "1";
                var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, style) }.ToList();
                //var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, "2") }.ToList();

                attributes.Add(new OpenXmlAttribute("t", null, "s"));//shared string type
                Writer.WriteStartElement(new Cell(), attributes);
                if (!ShareStringDictionary.ContainsKey(cells[i].Value))
                {
                    ShareStringDictionary.Add(cells[i].Value, ShareStringDictionary.Keys.Count());
                }

                //writing the index as the cell value
                Writer.WriteElement(new CellValue(ShareStringDictionary[cells[i].Value].ToString()));


                Writer.WriteEndElement();//cell

            }
            Writer.WriteEndElement(); //end of Row tag
        }

        public void Dispose()
        {
            Writer.WriteEndElement(); //end of SheetData
            Writer.WriteEndElement(); //end of worksheet
            Writer.Close();// TODO: poner en una clase sheet
            CreateShareStringPart();
            Writer.Dispose();
        }

        private void CreateShareStringPart()
        {
            if (ShareStringDictionary.Count() > 0)
            {
                var sharedStringPart = WorkbookPart.AddNewPart<SharedStringTablePart>();
                using (var writer = OpenXmlWriter.Create(sharedStringPart))
                {
                    writer.WriteStartElement(new SharedStringTable());
                    foreach (var item in ShareStringDictionary)
                    {
                        writer.WriteStartElement(new SharedStringItem());
                        writer.WriteElement(new Text(item.Key));
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }
            }

        }
    }
}
