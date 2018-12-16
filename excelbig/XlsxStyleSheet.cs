using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace excelbig
{
    public class XlsxStyleSheet
    {
        private readonly WorkbookPart WorkbookPart;
        private readonly Stylesheet Stylesheet;
        public readonly Dictionary<string, string> CellFormats;
        private uint IExcelIndex = 165; //number less than 164 is reserved by excel for default formats (number formats)
        public const string DefaultStyle = "Default";

        public XlsxStyleSheet(WorkbookPart workbookPart)
        {
            WorkbookPart = workbookPart;
            Stylesheet = CreateDefaultStylesheet();
            CellFormats = new Dictionary<string, string>()
            {
                { DefaultStyle, "0" }
            };
        }

        public UInt32Value AddDateFormat(string format)
        {
            var nfs = Stylesheet.NumberingFormats;            
            NumberingFormat nf;
            nf = new NumberingFormat();
            nf.NumberFormatId = IExcelIndex++;
            nf.FormatCode = format;
            nfs.Append(nf);
            nfs.Count = (uint)nfs.ChildElements.Count;
            return nf.NumberFormatId;
        }

        public void AddCellFormat(XlsxCellFormat format)
        {
            var fillNumber = 0U;
            if (format.ForegroundColor != null)
            {
                var fill = new Fill()
                {
                    PatternFill = new PatternFill()
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor()
                        {
                            Rgb = HexBinaryValue.FromString(format.ForegroundColor.HexaCode)
                        }
                    }
                };
                Stylesheet.Fills.Append(fill);
                Stylesheet.Fills.Count = (uint)Stylesheet.Fills.ChildElements.Count;
                fillNumber = Stylesheet.Fills.Count - 1;
            }
            var numberFormatId = 0U;
            if(!string.IsNullOrEmpty(format.NumberFormat))
            {
                numberFormatId = AddDateFormat(format.NumberFormat);
            }

            if(format.Border != null)
            {
                var leftColor = format.Border.Color != null ? new Color() { Rgb = HexBinaryValue.FromString(format.Border.Color.HexaCode) } : new Color() { Auto = true };
                var rightColor = format.Border.Color != null ? new Color() { Rgb = HexBinaryValue.FromString(format.Border.Color.HexaCode) } : new Color() { Auto = true };
                var topColor = format.Border.Color != null ? new Color() { Rgb = HexBinaryValue.FromString(format.Border.Color.HexaCode) } : new Color() { Auto = true };
                var bottomColor = format.Border.Color != null ? new Color() { Rgb = HexBinaryValue.FromString(format.Border.Color.HexaCode) } : new Color() { Auto = true };
                var diagonalColor = format.Border.Color != null ? new Color() { Rgb = HexBinaryValue.FromString(format.Border.Color.HexaCode) } : new Color() { Auto = true };
                
                Border border = new Border();
                border.LeftBorder = new LeftBorder(leftColor) { Style = format.Border.LeftBorder };
                border.RightBorder = new RightBorder(rightColor) { Style = format.Border.LeftBorder };
                border.TopBorder = new TopBorder(topColor) { Style = format.Border.LeftBorder };
                border.BottomBorder = new BottomBorder(bottomColor) { Style = format.Border.LeftBorder };
                border.DiagonalBorder = new DiagonalBorder(diagonalColor) { Style = format.Border.LeftBorder };
                Stylesheet.Borders.Append(border);
                Stylesheet.Borders.Count = (uint)Stylesheet.Borders.ChildElements.Count;
            }

            var cf = new CellFormat();
            cf.NumberFormatId = numberFormatId;
            cf.FontId = 0;
            cf.FillId = fillNumber;
            cf.ApplyFill = true;
            cf.BorderId = format.Border != null ? Stylesheet.Borders.Count - 1 : 0;
            cf.FormatId = 0;
            Stylesheet.CellFormats.Append(cf);

            if (CellFormats.ContainsKey(format.Name))
            {
                throw new Exception($"Formato duplicado {format.Name}");
            }
            Stylesheet.CellFormats.Count = (uint)Stylesheet.CellFormats.ChildElements.Count;
            CellFormats.Add(format.Name, $"{CellFormats.Count}"); //Tengo un estilo por default provisto por excel
        }

        public void Save()
        {
            var workbookStylesPart = WorkbookPart.AddNewPart<WorkbookStylesPart>();
            var style = workbookStylesPart.Stylesheet = Stylesheet;
            style.Save();
        }

        /// <summary>
        /// create the default excel formats.  These formats are required for the excel in order for it to render
        /// correctly.
        /// </summary>
        /// <returns></returns>
        private Stylesheet CreateDefaultStylesheet()
        {

            Stylesheet ss = new Stylesheet();

            Fonts fts = new Fonts();
            DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
            FontName ftn = new FontName();
            ftn.Val = "Calibri";
            FontSize ftsz = new FontSize();
            ftsz.Val = 11;
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.Append(ft);
            fts.Count = (uint)fts.ChildElements.Count;

            Fills fills = new Fills();
            Fill fill;
            PatternFill patternFill;

            //default fills used by Excel, don't changes these

            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.None;
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Gray125;
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);



            fills.Count = (uint)fills.ChildElements.Count;

            Borders borders = new Borders();
            Border border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.BottomBorder = new BottomBorder();
            border.DiagonalBorder = new DiagonalBorder();
            borders.Append(border);
            borders.Count = (uint)borders.ChildElements.Count;

            CellStyleFormats csfs = new CellStyleFormats();
            CellFormat cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            csfs.Append(cf);
            csfs.Count = (uint)csfs.ChildElements.Count;


            CellFormats cfs = new CellFormats();

            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cfs.Append(cf);



            var nfs = new NumberingFormats();



            nfs.Count = (uint)nfs.ChildElements.Count;
            cfs.Count = (uint)cfs.ChildElements.Count;

            ss.Append(nfs);
            ss.Append(fts);
            ss.Append(fills);
            ss.Append(borders);
            ss.Append(csfs);
            ss.Append(cfs);

            CellStyles css = new CellStyles(
                new CellStyle()
                {
                    Name = "Normal",
                    FormatId = 0,
                    BuiltinId = 0,
                }
                );

            css.Count = (uint)css.ChildElements.Count;
            ss.Append(css);

            DifferentialFormats dfs = new DifferentialFormats();
            dfs.Count = 0;
            ss.Append(dfs);

            TableStyles tss = new TableStyles();
            tss.Count = 0;
            tss.DefaultTableStyle = "TableStyleMedium9";
            tss.DefaultPivotStyle = "PivotStyleLight16";
            ss.Append(tss);
            return ss;
        }
    }
}
