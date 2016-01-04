using System;
using System.Collections.Generic;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Timer = System.Threading.Timer;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private static string wavPath = @"c:\1.wav";
        private static SoundPlayer p;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (p == null)
            {
                try
                {
                    p = new SoundPlayer(wavPath);
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Can't load " + wavPath + ": " + exception.Message);
                }
            }
            if (p != null)
                Application.NewMailEx += Application_NewMailEx;
        }

        void Application_NewMailEx(string anEntryId)
        {
            try
            {
                Outlook.NameSpace outlookNS = Application.GetNamespace("MAPI");
                Outlook.MAPIFolder folder = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Outlook.MailItem mailItem = (Outlook.MailItem)outlookNS.GetItemFromID(anEntryId, folder.StoreID);
                Timer t = new Timer(Bling2, mailItem, 2000, 0);
            }
            catch (Exception e)
            {
            }
        }

        private void Bling2(object state)
        {
            Bling((Outlook.MailItem)state);
        }

        private static void Bling(Outlook.MailItem mailItem)
        {
            try
            {
                if (mailItem.Subject.Contains("Boom"))
                {
                    //MessageBox.Show("Trying to play");
                    p.Play();
                }
            }
            catch (Exception e)
            {
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
