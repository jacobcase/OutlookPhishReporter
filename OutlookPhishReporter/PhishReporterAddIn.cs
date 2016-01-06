using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace OutlookPhishReporter
{
    public partial class PhishReporterAddIn
    {
        //Edit these 2 strings and 1 array to contain all the values you want. mailRecipients can have multiple
        //addresses if you want that for some reason.
        private string mailSubjectPrefix = "[PhishReport] ";
        private string[] mailRecipients = new string[] { "email@example.com" };
        private string mailBody = "This is a user-submitted report of a phish email";

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new PhishReporterRibbon(this);
        }


        private void PhishReporterAddIn_Startup(object sender, System.EventArgs e)
        {
            
        }

        private void PhishReporterAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        //This method gets the currently active window in Outlook, gets the selected items, iterates through
        //them and attaches them to be submitted as phishing.  This is where all the magic happens.
        public void SubmitSelectedEmails()
        {
            
            Outlook.Explorer activeExplorere = this.Application.ActiveExplorer();
            Outlook.Selection currentSelection = activeExplorere.Selection;

            DialogResult dialogResult = MessageBox.Show(
                "The selected emails will be submitted to your security team. Proceed?",
                "Report Phishing?",
                MessageBoxButtons.YesNo);

            if (currentSelection.Count == 0)
            {
                MessageBox.Show("Please select messages to report and try again");
                return;
            }


            if (dialogResult == DialogResult.Yes)
            {
                foreach (object selected in currentSelection)
                {
                    Outlook.MailItem mailItem;
                    //It may be possible that the selected item is not an email. I'm not sure under what
                    //conditions this could occure, but better to handle it just in case.
                    try
                    {
                        mailItem = (Outlook.MailItem)selected;
                    }
                    catch (InvalidCastException) { continue; }

                    
                    //An email is created that will be send to security and will have the phishing email attached.
                    Outlook.MailItem submission = Application.CreateItem(Outlook.OlItemType.olMailItem);

                    foreach (string recipient in mailRecipients)
                    {
                        submission.Recipients.Add(recipient);
                    }

                    submission.Subject = this.mailSubjectPrefix + mailItem.Subject;
                    submission.Attachments.Add(mailItem, Outlook.OlAttachmentType.olEmbeddeditem);
                    submission.Body = this.mailBody;
                    submission.Send();
                    mailItem.Delete();
                }
            }

            
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(PhishReporterAddIn_Startup);
            this.Shutdown += new System.EventHandler(PhishReporterAddIn_Shutdown);
        }
        
        #endregion
    }
}
