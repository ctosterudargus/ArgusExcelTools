using System.Collections.Generic;

namespace ArgusElectrical
{
    internal class TraceResult
    {
        public List<Cable> Cables { get; } = new();
        public List<Raceway> Raceways { get; } = new();
    }
}
