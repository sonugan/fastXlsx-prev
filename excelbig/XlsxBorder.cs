
using DocumentFormat.OpenXml.Spreadsheet;

namespace excelbig
{
    public class XlsxBorder
    {
        public XlsxBorder()
        {

        }

        public XlsxBorder(BorderStyleValues leftBorder = BorderStyleValues.None,
            BorderStyleValues rightBorder = BorderStyleValues.None,
            BorderStyleValues topBorder = BorderStyleValues.None,
            BorderStyleValues bottomBorder = BorderStyleValues.None,
            BorderStyleValues diagonalBorder = BorderStyleValues.None)
        {
            LeftBorder = leftBorder;
            RightBorder = rightBorder;
            TopBorder = topBorder;
            BottomBorder = bottomBorder;
            DiagonalBorder = diagonalBorder;
        }

        public static XlsxBorder CreateBox(BorderStyleValues style, XlsxColor color = null)
        {
            return new XlsxBorder()
            {
                LeftBorder = style,
                RightBorder = style,
                TopBorder = style,
                BottomBorder = style,
                Color = color
            };
        }

        public BorderStyleValues LeftBorder { get; set; }
        public BorderStyleValues RightBorder { get; set; }
        public BorderStyleValues TopBorder { get; set; }
        public BorderStyleValues BottomBorder { get; set; }
        public BorderStyleValues DiagonalBorder { get; set; }
        public XlsxColor Color { get; set; }
    }
}
