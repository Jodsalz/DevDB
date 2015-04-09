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

        static Outlook.MAPIFolder inboxFolder;
        static Outlook.MAPIFolder einbuchung;
        static Outlook._NameSpace outlookNS;
        static NetOffice.OutlookApi.Application outlookApplication;
        static int flag = 0;

        static void Main(string[] args)
        {

            // start outlook
            outlookApplication = new Outlook.Application();
            
            // get inbox
            outlookNS = outlookApplication.Session;
            inboxFolder = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["Inbox"];

            einbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Einbuchung"];
            Outlook.MAPIFolder ausbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Ausbuchung"];


            outlookApplication.NewMailExEvent += new Outlook.Application_NewMailExEventHandler(outlook_newmail);
            outlookApplication.Session.SyncObjects[1].SyncEndEvent += new Outlook.SyncObject_SyncEndEventHandler(sync_end);

            while (true)
            {
                
               
            }
            
        }

        private static void outlook_newmail(string s)
        {
            Console.WriteLine("enter newmail event");
            Console.WriteLine("starting sendandreceiv");
            outlookNS.SendAndReceive(false);

        }

        private static void sync_end()
        {
            Console.WriteLine("enter syncend event");

            string subject = "leer";
            string body = "leer";
            Outlook.MailItem mailItem = null;

            Console.WriteLine("Entered sync_end Event with inboxitems: " + inboxFolder.Items.Count);

            for (int i = inboxFolder.Items.Count; i > 0; i--)
            {
                mailItem = inboxFolder.Items[i] as Outlook.MailItem;

                if (mailItem.DownloadState != OlDownloadState.olFullItem)
                {
                    Console.WriteLine("again");
                    mailItem.MarkForDownload = OlRemoteStatus.olMarkedForDownload;
                    outlookNS.SendAndReceive(false);
                    break;
                }

                if (mailItem != null && mailItem.Subject.StartsWith("DevDB 1", System.StringComparison.CurrentCultureIgnoreCase) && mailItem.DownloadState == OlDownloadState.olFullItem)
                {
                    subject = mailItem.Subject;
                    body = mailItem.Body;
                    einbuchung_vorgang(subject, body);
                    mailItem.Delete();
                }
                else if (mailItem != null && mailItem.Subject.StartsWith("DevDB 2", System.StringComparison.CurrentCultureIgnoreCase) && mailItem.DownloadState == OlDownloadState.olFullItem)
                {
                    subject = mailItem.Subject;
                    body = mailItem.Body;
                    ausbuchung_vorgang(subject, body);
                    mailItem.Delete();
                }

            }
            Console.WriteLine("Left Event");
        }


        private static void einbuchung_vorgang(string s, string se)
        {
            Console.WriteLine("Einbuchung: " + s + ": " + se + " erfolgreich erledigt");
        }

        private static void ausbuchung_vorgang(string s, string se)
        {
            Console.WriteLine("Einbuchung: " + s + ": " + se + " erfolgreich erledigt");
        }
       
       
    }
}
