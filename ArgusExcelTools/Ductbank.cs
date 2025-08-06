using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArgusElectrical
{
    internal class Ductbank
    {
        public string ID { get; set; }
        public List<string> Raceways { get; set; }

        public Ductbank()
        {
            List<string> Raceways = new List<string>();
        }
    }
}
