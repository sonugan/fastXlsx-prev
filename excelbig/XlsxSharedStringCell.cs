using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelbig
{
    public class XlsxSharedStringCell : XlsxCell
    {
        public XlsxSharedStringCell(string value, string formatName = "")
            : base(formatName, value)
        {
        }
    }
}
