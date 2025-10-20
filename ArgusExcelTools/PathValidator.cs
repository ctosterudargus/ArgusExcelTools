using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ArgusExcelTools.Constraints;
using static ArgusExcelTools.RouteCheckResult;

namespace ArgusExcelTools
{
    internal class PathValidator
    {
        /// <summary>
        /// Strict declared-route contiguity check. Trims ends only.
        /// </summary>
        public static RouteCheckResult CheckDeclaredRoute(
    Graph g,
    Cable cable,
    IReadOnlyDictionary<string, Raceway> racewayById)
        {
            var result = new RouteCheckResult { CableId = cable.ID, Status = RouteStatus.Valid };

            // Strict endpoints (trim ends only)
            var from = (cable.From ?? string.Empty).Trim();
            var to = (cable.To ?? string.Empty).Trim();

            // Parse declared list (keep strict tokens)
            var declared = ParseDeclaredRacewayList(cable.RacewayRouting);

            // Nothing declared -> let reachability handle it later
            if (declared.Count == 0)
                return result;

            // Resolve first raceway
            if (!racewayById.TryGetValue(declared[0], out var first))
            {
                result.Status = RouteStatus.BrokenRoute;
                result.OffendingRacewayA = declared[0];
                result.Message = $"Declared route references missing raceway ID '{declared[0]}'.";
                return result;
            }

            // If the first item doesn't touch FROM but touches TO, we treat the list as reversed.
            bool firstTouchesFrom = Touches(from, first);
            bool firstTouchesTo = Touches(to, first);

            if (!firstTouchesFrom)
            {
                if (firstTouchesTo)
                {
                    declared.Reverse();
                    // re-fetch the new first after reversal
                    if (!racewayById.TryGetValue(declared[0], out first))
                    {
                        result.Status = RouteStatus.BrokenRoute;
                        result.OffendingRacewayA = declared[0];
                        result.Message = $"Declared route references missing raceway ID '{declared[0]}'.";
                        return result;
                    }
                }
                else
                {
                    result.Status = RouteStatus.BrokenRoute;
                    result.OffendingRacewayA = declared[0];
                    result.Message = $"Declared route does not start at '{from}' or end at '{to}'.";
                    return result;
                }
            }

            // Special case: single raceway declared -> must touch BOTH endpoints (order-insensitive)
            if (declared.Count == 1)
            {
                var touchesFrom = Touches(from, first);
                var touchesToEnd = Touches(to, first);
                if (!touchesFrom || !touchesToEnd)
                {
                    result.Status = RouteStatus.BrokenRoute;
                    result.OffendingRacewayA = first.ID;
                    result.Message = $"Declared route with a single raceway '{first.ID}' does not connect '{from}' to '{to}'.";
                }
                return result;
            }

            // Contiguity check for each consecutive pair
            for (int i = 0; i < declared.Count - 1; i++)
            {
                Raceway a = null, b = null;

                if (!racewayById.TryGetValue(declared[i], out a) ||
                    !racewayById.TryGetValue(declared[i + 1], out b))
                {
                    result.Status = RouteStatus.BrokenRoute;
                    result.OffendingRacewayA = a?.ID ?? declared[i];
                    result.OffendingRacewayB = b?.ID ?? declared[i + 1];
                    result.Message = "Declared route references missing raceway ID(s) in sequence.";
                    return result;
                }

                var common = CommonNode(a, b);
                if (common == null)
                {
                    result.Status = RouteStatus.BrokenRoute;
                    result.OffendingRacewayA = a.ID;
                    result.OffendingRacewayB = b.ID;
                    result.Message = $"Declared raceways {a.ID} and {b.ID} are not contiguous.";
                    return result;
                }
            }

            // Final check: last raceway must touch TO
            var lastId = declared[declared.Count - 1];
            if (!racewayById.TryGetValue(lastId, out var last) || !Touches(to, last))
            {
                result.Status = RouteStatus.BrokenRoute;
                result.OffendingRacewayA = lastId;
                result.Message = $"Declared route does not end at '{to}'.";
                return result;
            }

            // If we got here, declared route is contiguous and connects FROM -> TO (order-insensitive)
            return result;
        }


        /// <summary>
        /// Connectivity check with constraints. Returns true with path if reachable.
        /// </summary>
        public static bool TryFindPath(
            Graph g,
            string fromRaw,
            string toRaw,
            Cable cable,
            IPathConstraint constraint,              // may be null for no constraints
            out List<string> pathNodes,
            out string constraintReason)
        {
            pathNodes = null;
            constraintReason = null;

            var from = (fromRaw ?? string.Empty).Trim();
            var to = (toRaw ?? string.Empty).Trim();

            if (!g.Nodes.Contains(from) || !g.Nodes.Contains(to))
                return false;

            var q = new Queue<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var parent = new Dictionary<string, Tuple<string, GraphEdge>>(StringComparer.Ordinal);

            q.Enqueue(from);
            seen.Add(from);

            while (q.Count > 0)
            {
                var u = q.Dequeue();
                if (StringComparer.Ordinal.Equals(u, to))
                {
                    pathNodes = Reconstruct(parent, from, to);
                    return true;
                }

                foreach (var e in g.Adj[u])
                {
                    var v = e.Other(u);
                    if (seen.Contains(v))
                        continue;

                    if (constraint != null && !constraint.AllowsEdge(cable, e, out constraintReason))
                        continue;

                    seen.Add(v);
                    parent[v] = Tuple.Create(u, e);
                    q.Enqueue(v);
                }
            }

            return false;
        }

        // --- helpers ---

        private static List<string> ParseDeclaredRacewayList(string routing)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(routing)) return list;

            // Split on common delimiters; keep strict tokens
            var parts = routing.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var t = p.Trim();
                if (t.Length > 0) list.Add(t);
            }
            return list;
        }

        private static bool Touches(string node, Raceway r)
        {
            var a = (r.From ?? string.Empty).Trim();
            var b = (r.To ?? string.Empty).Trim();
            return StringComparer.Ordinal.Equals(node, a) || StringComparer.Ordinal.Equals(node, b);
        }

        private static string CommonNode(Raceway a, Raceway b)
        {
            var a1 = (a.From ?? string.Empty).Trim();
            var a2 = (a.To ?? string.Empty).Trim();
            var b1 = (b.From ?? string.Empty).Trim();
            var b2 = (b.To ?? string.Empty).Trim();

            if (StringComparer.Ordinal.Equals(a1, b1) || StringComparer.Ordinal.Equals(a1, b2)) return a1;
            if (StringComparer.Ordinal.Equals(a2, b1) || StringComparer.Ordinal.Equals(a2, b2)) return a2;
            return null;
        }

        private static List<string> Reconstruct(
            Dictionary<string, Tuple<string, GraphEdge>> parent,
            string from, string to)
        {
            var path = new List<string>();
            var cur = to;
            while (!StringComparer.Ordinal.Equals(cur, from))
            {
                path.Add(cur);
                cur = parent[cur].Item1;
            }
            path.Add(from);
            path.Reverse();
            return path;
        }
    }
}
