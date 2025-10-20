using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArgusExcelTools
{
    public partial class Argus
    {
        private TraceResult traceContext;
        private void Argus_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConduitFill_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnCableTrace_Click(object sender, RibbonControlEventArgs e)
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

        private void btnCreateSchedule_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            try
            { 
                var activeWb = app.ActiveWorkbook;
                if (activeWb != null)
                {
                    activeWb.Close();
                }
                string templatePath = TemplateHelper.ExtractTemplate("ArgusTemplate.xltm");
                app.Workbooks.Open(templatePath);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Failed to create workbook from template:\n{ex.Message}");
            }
        }


        private void btnCreateMechSchedule_Click(object sender, RibbonControlEventArgs e)
        {
            new ArgusExcelTools.MechScheduleBuilder().BuildFromActiveSheet();
        }

        private void btnUpdateMechSchedule_Click(object sender, RibbonControlEventArgs e)
        {
            
        }
    }
}
