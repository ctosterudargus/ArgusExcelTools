using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;

namespace ArgusElectrical
{
    public class ArgusRibbon : RibbonBase
    {
        private RibbonTab argusTab;
        private RibbonGroup toolsGroup;
        private RibbonButton cableTraceButton;
        private RibbonButton generateScheduleButton;

        private TraceResult traceContext;

        public ArgusRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            var factory = Globals.Factory.GetRibbonFactory();

            argusTab = factory.CreateRibbonTab();
            argusTab.Label = "Argus";

            toolsGroup = factory.CreateRibbonGroup();
            toolsGroup.Label = "Tools";

            cableTraceButton = factory.CreateRibbonButton();
            cableTraceButton.Label = "Cable Trace";
            cableTraceButton.Click += CableTraceButton_Click;

            generateScheduleButton = factory.CreateRibbonButton();
            generateScheduleButton.Label = "Generate Schedule";
            generateScheduleButton.Click += GenerateScheduleButton_Click;

            toolsGroup.Items.Add(cableTraceButton);
            toolsGroup.Items.Add(generateScheduleButton);
            argusTab.Groups.Add(toolsGroup);
            Tabs.Add(argusTab);
        }

        private void CableTraceButton_Click(object sender, RibbonControlEventArgs e)
        {
            var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            var processor = new CableTraceProcessor();

            try
            {
                traceContext = processor.Run(workbook);
                MessageBox.Show("Cable trace complete");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Cable Trace");
            }
        }

        private void GenerateScheduleButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (traceContext == null)
            {
                MessageBox.Show("Run Cable Trace before generating the schedule");
                return;
            }

            var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            var generator = new ScheduleGenerator();

            try
            {
                generator.Generate(workbook, traceContext);
                MessageBox.Show("Cable schedule generated");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Generate Schedule");
            }
        }
    }
}
