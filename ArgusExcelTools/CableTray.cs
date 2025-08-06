using System.Collections.Generic;

namespace ArgusExcelTools
{
    internal class CableTray
    {
        public string ID { get; set; }
        public List<string> Cables { get; }

        public CableTray()
        {
            Cables = new List<string>();
        }
    }
}
