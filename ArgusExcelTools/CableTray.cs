using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArgusElectrical
{
    internal class CableTray
    {
        public string ID { get; set; }
        public List<string> Cables { get; set; }

        public CableTray()
        {
            List<string> Cables = new List<string>();
        }
    }
}
