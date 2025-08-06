using System.Collections.Generic;

namespace ArgusExcelTools
{
    internal class Ductbank
    {
        public string ID { get; set; }
        public List<string> Raceways { get; }

        public Ductbank()
        {
            Raceways = new List<string>();
        }
    }
}
