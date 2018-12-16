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

        public XlsxWorksSheet(XlsxStyleSheet style, WorkbookPart workbookPart, WorksheetPart worksheetPart, IList<XlsxColumn> columns)
        {
            Style = style;
            WorkbookPart = workbookPart;
            Writer = OpenXmlWriter.Create(worksheetPart);
            Writer.WriteStartElement(new Worksheet());

            //TODO: definir las columnas INI
            WriteColumnDefinition(columns);
            //TODO: definir las columnas FIN

            Writer.WriteStartElement(new SheetData()); // TODO: poner en una clase sheet
        }

        private void WriteColumnDefinition(IList<XlsxColumn> columns)
        {
            var orderedColumns = columns
                .OrderBy(c => c.ColumnNumber)
                .Aggregate(new List<XlsxColumnGroup>(), (l, c) =>
                {
                    if (l.Any() && l.Last().Width == c.Width)
                    {
                        l.Last().MaxColumnNumber = c.ColumnNumber;
                    }
                    else
                    {
                        l.Add(new XlsxColumnGroup(c.ColumnNumber, c.Width));
                    }
                    return l;
                });

            Writer.WriteStartElement(new Columns());
            foreach (var column in orderedColumns)
            {
                //List<OpenXmlAttribute> attrs = new List<OpenXmlAttribute>();
                //attrs = new List<OpenXmlAttribute>();
                //// min and max are required attributes
                //// This means from columns 2 to 4, both inclusive
                //attrs.Add(new OpenXmlAttribute("min", null, "2"));
                //attrs.Add(new OpenXmlAttribute("max", null, "4"));
                //attrs.Add(new OpenXmlAttribute("width", null, "25"));
                //Writer.WriteStartElement(new Column(), attrs);
                //Writer.WriteEndElement();

                List<OpenXmlAttribute> attrs = new List<OpenXmlAttribute>();
                attrs = new List<OpenXmlAttribute>();
                attrs.Add(new OpenXmlAttribute("min", null, column.MinColumnNumber.ToString()));
                attrs.Add(new OpenXmlAttribute("max", null, column.MaxColumnNumber.ToString()));
                attrs.Add(new OpenXmlAttribute("width", null, column.Width.ToString()));
                Writer.WriteStartElement(new Column(), attrs);
                Writer.WriteEndElement();
            }
            Writer.WriteEndElement();
        }
        
        public void WriteRow(IList<XlsxCell> cells)
        {
            Writer.WriteStartElement(new Row());
            for (int i = 0; i < cells.Count(); i++)
            {
                switch(cells[i])
                {
                    case XlsxSharedStringCell sc:
                        WriteSharedString(sc);
                        break;
                    case XlsxDateCell dc:
                        WriteDate(dc);
                        break;
                    case XlsxNumberCell nc:
                        WriteNumber(nc);
                        break;
                }
            }
            Writer.WriteEndElement(); //end of Row tag
        }

        private void WriteSharedString(XlsxSharedStringCell cell)
        {
            var style = !string.IsNullOrEmpty(cell.FormatName) ? Style.CellFormats[cell.FormatName] : "0";
            var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, style) }.ToList();
            
            attributes.Add(new OpenXmlAttribute("t", null, "s"));//shared string type
            Writer.WriteStartElement(new Cell(), attributes);
            if (!ShareStringDictionary.ContainsKey(cell.Value))
            {
                ShareStringDictionary.Add(cell.Value, ShareStringDictionary.Keys.Count());
            }

            //writing the index as the cell value
            Writer.WriteElement(new CellValue(ShareStringDictionary[cell.Value].ToString()));
            
            Writer.WriteEndElement();//cell
        }

        private void WriteDate(XlsxDateCell cell)
        {
            var style = !string.IsNullOrEmpty(cell.FormatName) ? Style.CellFormats[cell.FormatName] : "0";
            var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, style) }.ToList();

            if (attributes == null)
            {
                Writer.WriteStartElement(new Cell() { DataType = CellValues.Number });
            }
            else
            {
                Writer.WriteStartElement(new Cell() { DataType = CellValues.Number }, attributes);
            }

            Writer.WriteElement(new CellValue(cell.Value));

            Writer.WriteEndElement();
        }

        private void WriteNumber(XlsxNumberCell cell)
        {
            var style = !string.IsNullOrEmpty(cell.FormatName) ? Style.CellFormats[cell.FormatName] : "0";
            var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, style) }.ToList();

            if (attributes == null)
            {
                Writer.WriteStartElement(new Cell() { DataType = CellValues.Number });
            }
            else
            {
                Writer.WriteStartElement(new Cell() { DataType = CellValues.Number }, attributes);
            }

            Writer.WriteElement(new CellValue(cell.Value));

            Writer.WriteEndElement();
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
