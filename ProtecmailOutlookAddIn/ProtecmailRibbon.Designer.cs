namespace ProtecmailOutlookAddIn
{
    partial class ProtecmailRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ProtecmailRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.reportSpam = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
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
            this.group1.Items.Add(this.reportSpam);
            this.group1.Name = "group1";
            // 
            // reportSpam
            // 
            this.reportSpam.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.reportSpam.Image = global::ProtecmailOutlookAddIn.Properties.Resources.protecmail_square;
            this.reportSpam.Label = "Report spams";
            this.reportSpam.Name = "reportSpam";
            this.reportSpam.ScreenTip = "Report spams to Protecmail";
            this.reportSpam.ShowImage = true;
            this.reportSpam.SuperTip = "Select spam and click this button to report them to Protecmail";
            this.reportSpam.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.reportSpam_Click);
            // 
            // ProtecmailRibbon
            // 
            this.Name = "ProtecmailRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ProtecmailRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton reportSpam;
    }

    partial class ThisRibbonCollection
    {
        internal ProtecmailRibbon ProtecmailRibbon
        {
            get { return this.GetRibbon<ProtecmailRibbon>(); }
        }
    }
}
