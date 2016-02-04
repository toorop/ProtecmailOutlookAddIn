using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Diagnostics;

namespace ProtecmailOutlookAddIn
{
    public partial class ThisAddIn
    {

        private ProtecmailOutlookAddIn.ProtecmailRibbon protecmailRibbon;

        private  RestSharp.RestClient protecmailAPIClient;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            protecmailRibbon = Globals.Ribbons.ProtecmailRibbon;
            protecmailRibbon.SetAddin(this);

            // REST client
            protecmailAPIClient = new RestSharp.RestClient("http://reports.protecmail.com");


        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Remarque : Outlook ne déclenche plus cet événement. Si du code
            //    doit s'exécuter à la fermeture d'Outlook, voir http://go.microsoft.com/fwlink/?LinkId=506785
        }


        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        // Méthode qui va reporter les mails selectionnés 
         public void ReportSpams()
         {
             Outlook.Explorer activeExplorer = this.Application.ActiveExplorer();
             Outlook.Selection selection = activeExplorer.Selection;

             // Si il n'y a rien de selectionné... on n'a rien à faire
             if (selection.Count == 0)
             {
                 MessageBox.Show("You have to select at least one message");
                 return;
             }

             foreach (object selected in selection)
             {
                 Outlook.MailItem mailItem;

                 // Il peut y avoir autre choise de seclectionné que des mails
                 // donc si selected ce n'est pas un mailItem on ne le traite pas
                 try
                 {
                     mailItem = (Outlook.MailItem)selected;
                 } catch (InvalidCastException) { continue;}

                 //MessageBox.Show(mailItem.Raw());

                var request = new RestSharp.RestRequest("aj/report", RestSharp.Method.POST);

                //request.AddBody(mailItem.Raw());
                request.AddParameter("text/plain", mailItem.Raw(), RestSharp.ParameterType.RequestBody);
                RestSharp.IRestResponse response = protecmailAPIClient.Execute(request);               
                Debug.WriteLine(response.ResponseStatus);
                Debug.WriteLine(response.StatusCode.ToString());
                Debug.WriteLine(response.Content);
            }
            // retour client
            if (selection.Count == 1)
            {
                MessageBox.Show("Spam has been successfully reported to Protecmail");
            } else
            {
                MessageBox.Show("Spams has been successfully reported to Protecmail");
            }
        }
    }


    // On ajoute la methode Raw a MailItem qui retourne la source compléte d'un mail
    public static class MailItemExtension
    {
        // headers
        private const string TransportMessageHeadersSchema = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        public static string Raw(this Outlook.MailItem mailItem)
        {
            return (string)mailItem.PropertyAccessor.GetProperty(TransportMessageHeadersSchema) + "\n\n" + mailItem.Body;
        }

    }

}
