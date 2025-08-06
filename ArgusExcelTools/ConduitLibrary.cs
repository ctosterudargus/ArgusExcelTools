using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    internal class ConduitLibrary
    {
        public Dictionary<string, double> HDPE { get; set; }

        public ConduitLibrary()
        {
            HDPE = new Dictionary<string, double>();
            HDPE.Add("3/4\"", .508);
            HDPE.Add("1\"", .832);
            HDPE.Add("1 1/4\"", 1.453);
            HDPE.Add("1 1/2\"", 1.986);
            HDPE.Add("2\"", 3.291);
            HDPE.Add("2 1/2\"", 4.695);
            HDPE.Add("3\"", 7.268);
            HDPE.Add("3 1/2\"", 9.737);
            HDPE.Add("4\"", 12.554);
            HDPE.Add("5\"", 19.761);
            HDPE.Add("6\"", 28.567);


        }

    }
}
