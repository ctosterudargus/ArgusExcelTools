using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    internal class CableLibrary
    {
        public Dictionary<string, double> THHN { get; set; }
        public Dictionary<string, double> XHHW { get; set; }
        public Dictionary<string, double> TCER { get; set; }
        public Dictionary <string, double> BELDEN { get; set; }
        public Dictionary<string, double> PVC { get; set; }
        public Dictionary<string, double> OSP { get; set; }





        

        public CableLibrary()
        {
            THHN = new Dictionary<string, double>();
            XHHW = new Dictionary<string, double>();
            TCER = new Dictionary<string, double>();
            BELDEN = new Dictionary<string, double>();
            PVC = new Dictionary<string, double>();
            OSP = new Dictionary<string, double>();




            OSP.Add("2-CT SM-FO", .402); //D-002-LN-8W-F02NS
            OSP.Add("18-CT SM-FO", .402);
            OSP.Add("12-CT SM-FO", .402);
            OSP.Add("2-CT SMFO", .402); //D-002-LN-8W-F02NS
            OSP.Add("18-CT SMFO", .402);
            OSP.Add("12-CT SMFO", .402);
            OSP.Add("CAT-6", .251);

            
            BELDEN.Add("BELDEN 8441", .194);
            BELDEN.Add("BELDEN 88760", .148);
            BELDEN.Add("BELDEN 8443", .172);

            THHN.Add("#14", .111);
            THHN.Add("#12", .130);
            THHN.Add("#10", .164);
            THHN.Add("#8", .216);
            THHN.Add("#6", .254);
            THHN.Add("#4", .324);
            THHN.Add("#3", .352);
            THHN.Add("#2", .384);
            THHN.Add("#1", .446);
            THHN.Add("#1/0", .486);
            THHN.Add("#2/0", .532);
            THHN.Add("#3/0", .584);
            THHN.Add("#4/0", .642);
            THHN.Add("#250 KCMIL", .711);
            THHN.Add("#300 KCMIL", .766);
            THHN.Add("#350 KCMIL", .817);
            THHN.Add("#400 KCMIL", .864);
            THHN.Add("#500 KCMIL", .949);
            THHN.Add("#600 KCMIL", 1.051);
            THHN.Add("#700 KCMIL", 1.122);

            XHHW.Add("#14", .140);
            XHHW.Add("#12", .152);
            XHHW.Add("#10", .176);
            XHHW.Add("#8", .236);
            XHHW.Add("#6", .274);
            XHHW.Add("#4", .322);
            XHHW.Add("#3", .350);
            XHHW.Add("#2", .382);
            XHHW.Add("#1", .442);
            XHHW.Add("#1/0", .482);
            XHHW.Add("#2/0", .528);
            XHHW.Add("#3/0", .58);
            XHHW.Add("#4/0", .638);
            XHHW.Add("#250 KCMIL", .705);
            XHHW.Add("#300 KCMIL", .76);
            XHHW.Add("#350 KCMIL", .811);
            XHHW.Add("#400 KCMIL", .858);
            XHHW.Add("#500 KCMIL", .943);
            XHHW.Add("#600 KCMIL", 1.053);
            XHHW.Add("#700 KCMIL", 1.124);

            TCER.Add("4/C #16 TSP", .52);
            TCER.Add("2/C #16 TSP", .35); // https://www.okonite.com/media//catalog/product/files/5-50.pdf
            TCER.Add("2/C #14", .39);
            TCER.Add("3/C #14", .41);
            TCER.Add("4/C #14", .44);
            TCER.Add("5/C #14", .48);
            TCER.Add("7/C #14", .52);
            TCER.Add("9/C #14", .64);
            TCER.Add("12/C #14", .71);
            TCER.Add("19/C #14", .83);
            TCER.Add("37/C #14", 1.14);
            TCER.Add("2/C #12", .42);
            TCER.Add("3/C #12", .45);
            TCER.Add("4/C #12", .49);
            TCER.Add("5/C #12", .53);
            TCER.Add("7/C #12", .61);
            TCER.Add("9/C #12", .70);
            TCER.Add("12/C #12", .79);
            TCER.Add("19/C #12", .96);
            TCER.Add("37/C #12", 1.27);
            TCER.Add("2/C #10", .47);
            TCER.Add("3/C #10", .50);
            TCER.Add("4/C #10", .58);
            TCER.Add("5/C #10", .63);
            TCER.Add("7/C #10", .68);
            TCER.Add("9/C #10", .79);
            TCER.Add("12/C #10", .93);
            TCER.Add("3/C #8", .68);
            TCER.Add("4/C #8", .71);
            TCER.Add("3/C #6", .73);
            TCER.Add("4/C #6", .82);
            TCER.Add("3/C #4", .82);
            TCER.Add("4/C #4", .95);
            TCER.Add("3/C #2", .98);
            TCER.Add("4/C #2", 1.08);
            TCER.Add("3/C #1", 1.09);
            TCER.Add("4/C #1", 1.21);
            TCER.Add("3/C #1/0", 1.17);
            TCER.Add("4/C #1/0", 1.29);
            TCER.Add("3/C #2/0", 1.26);
            TCER.Add("4/C #2/0", 1.39);
            TCER.Add("3/C #4/0", 1.49);
            TCER.Add("4/C #4/0", 1.63);
            TCER.Add("3/C 250 KCMIL", 1.62);
            TCER.Add("4/C 250 KCMIL", 1.86);
            TCER.Add("3/C 350 KCMIL", 1.89);
            TCER.Add("4/C 350 KCMIL", 2.08);
            TCER.Add("3/C 500 KCMIL", 2.14);
            TCER.Add("4/C 500 KCMIL", 2.37);
            TCER.Add("3/C 750 KCMIL", 2.57);
            TCER.Add("4/C 750 KCMIL", 2.91);
            





        }
    }
}
