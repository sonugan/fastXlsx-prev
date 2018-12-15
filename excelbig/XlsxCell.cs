
namespace excelbig
{
    public class XlsxCell
    {
        public string FormatName { get; set; }
        public string Value { get; set; }

        public XlsxCell(string value, string formatName = "")
        {
            FormatName = formatName;
            Value = value;
        }
    }
}
