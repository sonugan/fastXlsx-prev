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

        public XlsxStyleSheet(WorkbookPart workbookPart)
        {
            WorkbookPart = workbookPart;
            Stylesheet = CreateDefaultStylesheet();
            CellFormats = new Dictionary<string, string>();
        }

        public void AddCellFormat(XlsxCellFormat format)
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

            var cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = Stylesheet.Fills.Count - 1;
            cf.ApplyFill = true;
            cf.BorderId = 0;
            cf.FormatId = 0;
            Stylesheet.CellFormats.Append(cf);

            if (CellFormats.ContainsKey(format.Name))
            {
                throw new Exception($"Formato duplicado {format.Name}");
            }
            Stylesheet.CellFormats.Count = (uint)Stylesheet.CellFormats.ChildElements.Count;
            CellFormats.Add(format.Name, $"{CellFormats.Count + 1}"); //Tengo un estilo por default provisto por excel
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
