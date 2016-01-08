using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.Drawing;
using System.Resources;

namespace OutlookPhishReporter
{
    [ComVisible(true)]
    public class PhishReporterRibbon : Office.IRibbonExtensibility
    {
        private string phishButtonLabel = "Report a Phish!";
        private string phishScreentip = "Report phishing emails";
        private string phishSupertip = "Click this button to report the selected emails to the security team for phishing";
        private string phishGroupLabel = "CIRT";

        private Office.IRibbonUI ribbon;
        private PhishReporterAddIn addin;

        public PhishReporterRibbon(PhishReporterAddIn addin)
        {
            this.addin = addin;
        }

        //This method returns a text string of the xml for the ribbon icon which is the 2 xml files in the
        //project.  Since when using the explorer and when viewing individual messages uses 2 different ribbons,
        //to get the plugin to show up in both windows, you need to have a version for each ribbon.  The XML
        //is nearly identical except for the idMso attribute of the <tab>.  This just identifies which window is 
        //requesting the UI string and returns the approriate XML, or null if it's a non-applicable window.
        public string GetCustomUI(string ribbonID)
        {          
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Mail.Read":
                    return GetResourceText("OutlookPhishReporter.PhishReporterRibbon.ReadMessage.xml");
                case "Microsoft.Outlook.Explorer":
                    return GetResourceText("OutlookPhishReporter.PhishReporterRibbon.Explorer.xml");
                default:
                    return null;                    
            }
        }

        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void OnPhishClick(Office.IRibbonControl control)
        {
            addin.SubmitSelectedEmails();
        }

        public Bitmap OnPhishGetImage(Office.IRibbonControl control)
        {
            return Resources.PhishingIcon;
        }

        /*
        I put all of these in callbacks rather than directly in the XML since
        you will probably want the same string for all of the ribbons the button
        will be on.
        */
        public string OnPhishLabel(Office.IRibbonControl control)
        {
            return this.phishButtonLabel;
        }

        public string OnPhishScreentip(Office.IRibbonControl control)
        {
            return this.phishScreentip;
        }

        public string OnPhishSupertip(Office.IRibbonControl control)
        {
            return this.phishSupertip;
        }

        public string OnGroupLabel(Office.IRibbonControl control)
        {
            return this.phishGroupLabel;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }


        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

    }
}
