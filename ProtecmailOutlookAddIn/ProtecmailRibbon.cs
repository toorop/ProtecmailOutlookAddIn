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
        private void ProtecmailRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void reportSpam_Click(object sender, RibbonControlEventArgs e)
        {
            Debug.WriteLine("bouton report spam clicked");
        }
    }
}
