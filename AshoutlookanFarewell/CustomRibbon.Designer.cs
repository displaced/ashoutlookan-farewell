
namespace AshoutlookanFarewell
{
    partial class CustomRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CustomRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CustomRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.buttonGroup = this.Factory.CreateRibbonGroup();
            this.sbtnNewMissive = this.Factory.CreateRibbonSplitButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.buttonGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.buttonGroup);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // buttonGroup
            // 
            this.buttonGroup.Items.Add(this.sbtnNewMissive);
            this.buttonGroup.Label = "Ashokan Farewell";
            this.buttonGroup.Name = "buttonGroup";
            // 
            // sbtnNewMissive
            // 
            this.sbtnNewMissive.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.sbtnNewMissive.Image = ((System.Drawing.Image)(resources.GetObject("sbtnNewMissive.Image")));
            this.sbtnNewMissive.Items.Add(this.button1);
            this.sbtnNewMissive.Label = "New Missive";
            this.sbtnNewMissive.Name = "sbtnNewMissive";
            this.sbtnNewMissive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sbtnNewMissive_Click);
            // 
            // button1
            // 
            this.button1.Label = "Customise...";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // CustomRibbon
            // 
            this.Name = "CustomRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CustomRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.buttonGroup.ResumeLayout(false);
            this.buttonGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup buttonGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton sbtnNewMissive;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal CustomRibbon CustomRibbon
        {
            get { return this.GetRibbon<CustomRibbon>(); }
        }
    }
}
