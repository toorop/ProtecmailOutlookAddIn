using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;

namespace ProtecmailOutlookAddIn
{
    public partial class ProtecmailRibbon
    {
        private ThisAddIn addin;
        private void ProtecmailRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        // Set addin 
        public void SetAddin(ThisAddIn addin)
        {
            this.addin = addin;
        }


        private void reportSpam_Click(object sender, RibbonControlEventArgs e)
        {
            addin.ReportSpams();
        }
    }
}
