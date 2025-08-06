namespace ArgusExcelTools
{
    partial class Argus : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Argus()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.argusTab = this.Factory.CreateRibbonTab();
            this.argusElectrical = this.Factory.CreateRibbonGroup();
            this.btnConduitFill = this.Factory.CreateRibbonButton();
            this.btnCableTrace = this.Factory.CreateRibbonButton();
            this.btnCreateSchedule = this.Factory.CreateRibbonButton();
            this.argusMechanical = this.Factory.CreateRibbonGroup();
            this.btnCreateMechSchedule = this.Factory.CreateRibbonButton();
            this.btnUpdateMechSchedule = this.Factory.CreateRibbonButton();
            this.argusCivil = this.Factory.CreateRibbonGroup();
            this.tab1.SuspendLayout();
            this.argusTab.SuspendLayout();
            this.argusElectrical.SuspendLayout();
            this.argusMechanical.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // argusTab
            // 
            this.argusTab.Groups.Add(this.argusElectrical);
            this.argusTab.Groups.Add(this.argusMechanical);
            this.argusTab.Groups.Add(this.argusCivil);
            this.argusTab.Label = "Argus";
            this.argusTab.Name = "argusTab";
            // 
            // argusElectrical
            // 
            this.argusElectrical.Items.Add(this.btnConduitFill);
            this.argusElectrical.Items.Add(this.btnCableTrace);
            this.argusElectrical.Items.Add(this.btnCreateSchedule);
            this.argusElectrical.Label = "Electrical";
            this.argusElectrical.Name = "argusElectrical";
            // 
            // btnConduitFill
            // 
            this.btnConduitFill.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConduitFill.Image = global::ArgusExcelTools.Properties.Resources.Conduit_Fill;
            this.btnConduitFill.Label = "Conduit Fill";
            this.btnConduitFill.Name = "btnConduitFill";
            this.btnConduitFill.ShowImage = true;
            this.btnConduitFill.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConduitFill_Click);
            // 
            // btnCableTrace
            // 
            this.btnCableTrace.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCableTrace.Image = global::ArgusExcelTools.Properties.Resources.CableTrace1;
            this.btnCableTrace.Label = "Cable Trace";
            this.btnCableTrace.Name = "btnCableTrace";
            this.btnCableTrace.ShowImage = true;
            this.btnCableTrace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCableTrace_Click);
            // 
            // btnCreateSchedule
            // 
            this.btnCreateSchedule.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCreateSchedule.Image = global::ArgusExcelTools.Properties.Resources.Create_Schedule;
            this.btnCreateSchedule.Label = "Create Schedule";
            this.btnCreateSchedule.Name = "btnCreateSchedule";
            this.btnCreateSchedule.ShowImage = true;
            this.btnCreateSchedule.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateSchedule_Click);
            // 
            // argusMechanical
            // 
            this.argusMechanical.Items.Add(this.btnCreateMechSchedule);
            this.argusMechanical.Items.Add(this.btnUpdateMechSchedule);
            this.argusMechanical.Label = "Mechanical";
            this.argusMechanical.Name = "argusMechanical";
            // 
            // btnCreateMechSchedule
            // 
            this.btnCreateMechSchedule.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCreateMechSchedule.Image = global::ArgusExcelTools.Properties.Resources.Mech_Valve_Schedule;
            this.btnCreateMechSchedule.Label = "Create Schedule";
            this.btnCreateMechSchedule.Name = "btnCreateMechSchedule";
            this.btnCreateMechSchedule.ShowImage = true;
            this.btnCreateMechSchedule.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateMechSchedule_Click);
            // 
            // btnUpdateMechSchedule
            // 
            this.btnUpdateMechSchedule.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateMechSchedule.Image = global::ArgusExcelTools.Properties.Resources.Update_Valve_Schedule;
            this.btnUpdateMechSchedule.Label = "Update Schedule";
            this.btnUpdateMechSchedule.Name = "btnUpdateMechSchedule";
            this.btnUpdateMechSchedule.ShowImage = true;
            this.btnUpdateMechSchedule.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateMechSchedule_Click);
            // 
            // argusCivil
            // 
            this.argusCivil.Label = "Civil";
            this.argusCivil.Name = "argusCivil";
            // 
            // Argus
            // 
            this.Name = "Argus";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.argusTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Argus_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.argusTab.ResumeLayout(false);
            this.argusTab.PerformLayout();
            this.argusElectrical.ResumeLayout(false);
            this.argusElectrical.PerformLayout();
            this.argusMechanical.ResumeLayout(false);
            this.argusMechanical.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab argusTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup argusElectrical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConduitFill;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCableTrace;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateSchedule;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup argusMechanical;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup argusCivil;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateMechSchedule;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateMechSchedule;
    }

    partial class ThisRibbonCollection
    {
        internal Argus Argus
        {
            get { return this.GetRibbon<Argus>(); }
        }
    }
}
