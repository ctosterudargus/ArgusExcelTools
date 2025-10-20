using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    internal class Graph
    {
        public HashSet<string> Nodes { get; } = new HashSet<string>(StringComparer.Ordinal);
        // Adjacency: node -> list of edges incident to node
        public Dictionary<string, List<GraphEdge>> Adj { get; } =
            new Dictionary<string, List<GraphEdge>>(StringComparer.Ordinal);

        public void AddNode(string nodeId)
        {
            if (!Adj.ContainsKey(nodeId))
            {
                Adj[nodeId] = new List<GraphEdge>();
                Nodes.Add(nodeId);
            }
        }

        public void AddEdge(GraphEdge e)
        {
            AddNode(e.U);
            AddNode(e.V);
            Adj[e.U].Add(e);
            Adj[e.V].Add(e);
        }
    }

    /// <summary>Undirected edge for a raceway segment.</summary>
    internal sealed class GraphEdge
    {
        public string RacewayId { get; }
        public string U { get; }     // endpoint A
        public string V { get; }     // endpoint B

        // Future-proof attributes (optional at first; populate later as data becomes available)
        public string RacewayType { get; }      // e.g., CONDUIT, TRAY, DUCTBANK
        public string CircuitType { get; }      // from schedule (power/control/etc.)
        public string Size { get; }             // e.g., "2\"", "3\"", "12x4 tray"
        public double? FillCapacityArea { get; } // sq in or mm^2 (capacity)
        public string ZoneClass { get; }        // hazardous/classification if applicable
        public double? Length { get; }          // if/when you have it

        public GraphEdge(
            string racewayId, string u, string v,
            string racewayType = null, string circuitType = null, string size = null,
            double? fillCapacityArea = null, string zoneClass = null, double? length = null)
        {
            RacewayId = racewayId;
            U = u;
            V = v;
            RacewayType = racewayType;
            CircuitType = circuitType;
            Size = size;
            FillCapacityArea = fillCapacityArea;
            ZoneClass = zoneClass;
            Length = length;
        }

        public string Other(string node) => StringComparer.Ordinal.Equals(node, U) ? V : U;
    }
}

