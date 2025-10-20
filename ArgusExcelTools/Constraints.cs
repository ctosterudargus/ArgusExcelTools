using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    internal class Constraints
    {
        internal interface IPathConstraint
        {
            /// <summary>
            /// Return true if this edge is traversable by the given cable. Reason is set on failure.
            /// </summary>
            bool AllowsEdge(Cable cable, GraphEdge edge, out string reason);
        }

        /// <summary>Compose multiple constraints (AND semantics).</summary>
        internal sealed class CompositeConstraint : IPathConstraint
        {
            private readonly IList<IPathConstraint> _rules;
            public CompositeConstraint(params IPathConstraint[] rules) => _rules = rules ?? Array.Empty<IPathConstraint>();

            public bool AllowsEdge(Cable cable, GraphEdge edge, out string reason)
            {
                foreach (var r in _rules)
                {
                    if (!r.AllowsEdge(cable, edge, out reason))
                        return false;
                }
                reason = null;
                return true;
            }
        }

        // === Examples you can implement when ready ===

        /// <summary>Signal class/type compatibility between cable and raceway circuit type.</summary>
        internal sealed class SignalTypeConstraint : IPathConstraint
        {
            public bool AllowsEdge(Cable cable, GraphEdge edge, out string reason)
            {
                reason = null;
                // Example placeholder: enforce if you have concrete mapping rules
                // if (!IsCompatible(cable.SignalType, edge.CircuitType)) { reason = "Incompatible signal/circuit type."; return false; }
                return true;
            }
        }

        /// <summary>Fill capacity check using area accounting (future: compute FillUsed elsewhere and pass in).</summary>
        internal sealed class FillCapacityConstraint : IPathConstraint
        {
            public bool AllowsEdge(Cable cable, GraphEdge edge, out string reason)
            {
                reason = null;
                // When you track per-raceway fill, compare required area vs remaining capacity here.
                return true;
            }
        }

        /// <summary>Zone/segregation constraint (Class/Div, etc.).</summary>
        internal sealed class ZoneSegregationConstraint : IPathConstraint
        {
            public bool AllowsEdge(Cable cable, GraphEdge edge, out string reason)
            {
                reason = null;
                return true;
            }
        }
    }
}
