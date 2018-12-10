using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WMSInterface
{
    class WMSLPNHeader
    {
        public string organizationCode;
        public string workOrder;
        public string moNumber;
        public string sourceCode;
        public string LPNQty;
        public string action;
        public List<WMSLPNItem> LPNItems;
    }
}
