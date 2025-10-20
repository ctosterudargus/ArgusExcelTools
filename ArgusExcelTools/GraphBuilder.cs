using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    internal class GraphBuilder
    {
        /// <summary>
        /// Build an undirected graph from raceways. Only trims ends; exact match otherwise.
        /// </summary>
        public static Graph Build(IEnumerable<Raceway> raceways)
        {
            var g = new Graph();
            foreach (var r in raceways)
            {
                // Trim ends only (strict rule)
                var u = (r.From ?? string.Empty).Trim();
                var v = (r.To ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(u) || string.IsNullOrEmpty(v))
                    continue; // QC issue; you can log later

                // You can compute FillCapacityArea later if/when you have geometry
                var edge = new GraphEdge(
                    racewayId: r.ID,
                    u: u,
                    v: v,
                    racewayType: null,         // plug in when ready
                    circuitType: r.CircuitType,
                    size: r.Size,
                    fillCapacityArea: null,
                    zoneClass: null,
                    length: null
                );

                g.AddEdge(edge);
            }
            return g;
        }
    }
}
