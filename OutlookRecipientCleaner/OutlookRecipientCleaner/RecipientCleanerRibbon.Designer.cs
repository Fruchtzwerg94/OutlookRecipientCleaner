namespace OutlookRecipientCleaner
{
    partial class RecipientCleanerRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RecipientCleanerRibbon()
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
            this.group_RecipientCleaner = this.Factory.CreateRibbonGroup();
            this.button_About = this.Factory.CreateRibbonButton();
            this.splitButton_Clean = this.Factory.CreateRibbonSplitButton();
            this.tab1.SuspendLayout();
            this.group_RecipientCleaner.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group_RecipientCleaner);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group_RecipientCleaner
            // 
            this.group_RecipientCleaner.Items.Add(this.splitButton_Clean);
            this.group_RecipientCleaner.Label = "Recipient Cleaner";
            this.group_RecipientCleaner.Name = "group_RecipientCleaner";
            // 
            // button_About
            // 
            this.button_About.Label = "About";
            this.button_About.Name = "button_About";
            this.button_About.ShowImage = true;
            this.button_About.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_About_Click);
            // 
            // splitButton_Clean
            // 
            this.splitButton_Clean.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton_Clean.Description = "Clean recipients";
            this.splitButton_Clean.Image = global::OutlookRecipientCleaner.Properties.Resources.Clean;
            this.splitButton_Clean.Items.Add(this.button_About);
            this.splitButton_Clean.Label = "Clean";
            this.splitButton_Clean.Name = "splitButton_Clean";
            this.splitButton_Clean.ScreenTip = "Clean recipients";
            this.splitButton_Clean.SuperTip = "Removes all recipients, which are addressed multiple times";
            this.splitButton_Clean.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SplitButton_Clean_Click);
            // 
            // RecipientCleanerRibbon
            // 
            this.Name = "RecipientCleanerRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RecipientCleanerRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group_RecipientCleaner.ResumeLayout(false);
            this.group_RecipientCleaner.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_RecipientCleaner;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_About;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton_Clean;
    }

    partial class ThisRibbonCollection
    {
        internal RecipientCleanerRibbon RecipientCleanerRibbon
        {
            get { return this.GetRibbon<RecipientCleanerRibbon>(); }
        }
    }
}
