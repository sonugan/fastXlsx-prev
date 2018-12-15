using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelbig
{
    public class XlsxColor
    {
        public string HexaCode { get; set; }

        public XlsxColor(string hexaCode)
        {
            HexaCode = hexaCode;
        }

        //public static String HexConverter(Color c)
        //{
        //    return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        //}
    }
}
