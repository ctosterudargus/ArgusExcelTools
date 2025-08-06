using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    internal class Vessel
    {
        public Dictionary<string, string> Fields { get; set; } = new Dictionary<string, string>();

        public string vesselTag { get; set; }

        public string drawingNumber { get; set; }

        public string vesselType { get; set; }

        public string modelnumber { get; set; }

        public string volume { get; set; }

        public string flowRate { get; set; }

        public string pressure { get; set; }


    }
}
