﻿
namespace BankingQRCodeForExcel
{
    partial class RibbonButton : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonButton()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonButton));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Button = this.Factory.CreateRibbonGroup();
            this.buttonSwiss = this.Factory.CreateRibbonButton();
            this.buttonSEPA = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Button.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.Button);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // Button
            // 
            this.Button.Items.Add(this.buttonSwiss);
            this.Button.Items.Add(this.buttonSEPA);
            this.Button.Label = "Banking QRCode";
            this.Button.Name = "Button";
            // 
            // buttonSwiss
            // 
            this.buttonSwiss.Image = ((System.Drawing.Image)(resources.GetObject("buttonSwiss.Image")));
            this.buttonSwiss.Label = "Generate Swiss";
            this.buttonSwiss.Name = "buttonSwiss";
            this.buttonSwiss.ShowImage = true;
            this.buttonSwiss.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSwiss_Click);
            // 
            // buttonSEPA
            // 
            this.buttonSEPA.Image = ((System.Drawing.Image)(resources.GetObject("buttonSEPA.Image")));
            this.buttonSEPA.Label = "Generate SEPA";
            this.buttonSEPA.Name = "buttonSEPA";
            this.buttonSEPA.ShowImage = true;
            this.buttonSEPA.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSEPA_Click);
            // 
            // RibbonButton
            // 
            this.Name = "RibbonButton";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonButton_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Button.ResumeLayout(false);
            this.Button.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSwiss;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSEPA;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonButton RibbonButton
        {
            get { return this.GetRibbon<RibbonButton>(); }
        }
    }
}
