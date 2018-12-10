using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WMSInterface
{
    class WMSLPNItem
    {
        public string LPN;
        public string packingName;
        public string LPNSize;
        public string setNo;
        public string module;
        public string cartonNo;
        public string remark;
        public List<WMSLPNSnItem> SNItems;
    }
}
