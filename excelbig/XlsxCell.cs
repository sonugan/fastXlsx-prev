
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

    //Boolean Boolean.When the item is serialized out as xml, its value is "b".
    //Number Number. When the item is serialized out as xml, its value is "n".
    //Error Error. When the item is serialized out as xml, its value is "e".
    //SharedString Shared String.When the item is serialized out as xml, its value is "s".
    //String String. When the item is serialized out as xml, its value is "str".
    //InlineString Inline String.When the item is serialized out as xml, its value is "inlineStr".
    //Date d. When the item is serialized out as xml, its value is "d".This item is only available in Office2010.
}
