using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    internal class RouteCheckResult
    {
        internal enum RouteStatus
        {
            Valid,
            BrokenRoute,    // declared route not contiguous
            NotReachable,   // graph has no path between endpoints
            ViolatesConstraints
        }

            public string CableId { get; set; }
            public RouteStatus Status { get; set; }
            public string Message { get; set; }                 // concise summary
            public List<string> PathNodes { get; set; }         // optional path from BFS
            public string OffendingRacewayA { get; set; }       // for BrokenRoute
            public string OffendingRacewayB { get; set; }       // for BrokenRoute
            public string ConstraintReason { get; set; }        // first constraint failure reason (if any)
        
    }
}
