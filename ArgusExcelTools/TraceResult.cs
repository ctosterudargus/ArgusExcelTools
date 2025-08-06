using System.Collections.Generic;

namespace ArgusExcelTools
{
    internal class TraceResult
    {
        public List<Cable> Cables { get; } = new List<Cable>();
        public List<Raceway> Raceways { get; } = new List<Raceway>();
    }
}
