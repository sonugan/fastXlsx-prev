

namespace excelbig
{
    public class XlsxCellFormat
    {
        public string Name { get; set; }
        public XlsxColor ForegroundColor { get; set; }
        public XlsxBorder Border { get; set; } 
        public string NumberFormat { get; set; }

        public XlsxCellFormat()
        {

        }

        public XlsxCellFormat(string name, XlsxCellFormat baseFormat)
        {
            Name = name;
            ForegroundColor = baseFormat.ForegroundColor;
            Border = baseFormat.Border;
            NumberFormat = baseFormat.NumberFormat;
        }
    }
}
