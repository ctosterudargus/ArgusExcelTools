using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ArgusExcelTools
{
    public partial class Argus
    {
        private void Argus_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConduitFill_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnCableTrace_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnCreateSchedule_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                

                // Create helper and builder
                var helper = new CreateScheduleHelper(Globals.ThisAddIn.Application);
                var cableBuilder = new CableScheduleBuilder(Globals.ThisAddIn.Application, Globals.ThisAddIn.Application.ActiveWorkbook);
                var racewayBuilder = new RacewayScheduleBuilder(Globals.ThisAddIn.Application, Globals.ThisAddIn.Application.ActiveWorkbook);
                var ductbankBuilder = new DuctbankScheduleBuilder(Globals.ThisAddIn.Application, Globals.ThisAddIn.Application.ActiveWorkbook);
                var cableTrayBuilder = new CableTrayScheduleBuilder(Globals.ThisAddIn.Application, Globals.ThisAddIn.Application.ActiveWorkbook);


                cableBuilder.Build();
                racewayBuilder.Build();
                ductbankBuilder.Build();
                cableTrayBuilder.Build();

                System.Windows.Forms.MessageBox.Show("Schedule created successfully!", "Success", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }


        private void btnCreateMechSchedule_Click(object sender, RibbonControlEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv",
                Title = "Select Mechanical CSV File"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string csvFilePath = dialog.FileName;
                var processor = new MechCsvProcessor(Globals.ThisAddIn.Application);
                processor.ProcessMechanicalCsv(csvFilePath);
            }
        }

        private void btnUpdateMechSchedule_Click(object sender, RibbonControlEventArgs e)
        {
            
        }
    }
}
