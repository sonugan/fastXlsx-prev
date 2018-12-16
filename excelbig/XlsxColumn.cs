
namespace excelbig
{
    public class XlsxColumn
    {
        public int ColumnNumber { get; set; }
        public double Width { get; set; }
    }

    public class XlsxColumnGroup
    {
        public XlsxColumnGroup(int minColumnNumber, int maxColumnNumber, double width)
        {
            MinColumnNumber = minColumnNumber;
            MaxColumnNumber = maxColumnNumber;
            Width = width;
        }

        public XlsxColumnGroup(int columnNumber, double width)
        {
            MinColumnNumber = columnNumber;
            MaxColumnNumber = columnNumber;
            Width = width;
        }

        public int MinColumnNumber { get; set; }
        public int MaxColumnNumber { get; set; }
        public double Width { get; set; }
    }
}
