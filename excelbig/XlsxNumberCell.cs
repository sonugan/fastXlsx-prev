namespace excelbig
{
    public class XlsxNumberCell : XlsxCell
    {
        public XlsxNumberCell(double value, string formatName = "")
            : base(value.ToString(), formatName)
        {
        }
    }
}
