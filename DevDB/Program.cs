using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice.OutlookApi.Enums;

using NetOffice;
using Outlook = NetOffice.OutlookApi;

namespace DevDB
{
    class Program
    {


        static void Main(string[] args)
        {
            string body = "leer";

            // start outlook
            NetOffice.OutlookApi.Application outlookApplication = new Outlook.Application();
            


            // get inbox
            Outlook._NameSpace outlookNS = outlookApplication.Session;
            Outlook.MAPIFolder inboxFolder = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["Inbox"];
            /*
            Outlook.MAPIFolder einbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Einbuchung"];
            Outlook.MAPIFolder ausbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Ausbuchung"];
            */
            Outlook.MAPIFolder einbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Einbuchung"];
            Outlook.MAPIFolder ausbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Ausbuchung"];

            outlookApplication.NewMailExEvent += new Outlook.Application_NewMailExEventHandler(outlook_newmail);
            
        }

        private static void outlook_newmail(string s)
        {
            foreach (COMObject item in inboxFolder.Items)
            {
                Outlook.MailItem mailItem = item as Outlook.MailItem;
                if (mailItem != null && mailItem.Subject == "DevDB Einbuchung")
                {
                    mailItem.Move(einbuchung);
                    body = mailItem.Body;
                }
                else if (mailItem != null && mailItem.Subject == "DevDB Ausbuchung")
                {
                    mailItem.Move(ausbuchung);
                    body = mailItem.Body;
                }
            }
        }
    }
}
