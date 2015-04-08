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

        static void Main(string[] args)
        {

            // start outlook
            NetOffice.OutlookApi.Application outlookApplication = new Outlook.Application();
            


            // get inbox
            Outlook._NameSpace outlookNS = outlookApplication.Session;
            inboxFolder = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["Posteingang"];

            einbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Einbuchung"];
            Outlook.MAPIFolder ausbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Ausbuchung"];

            outlookApplication.NewMailExEvent += new Outlook.Application_NewMailExEventHandler(outlook_newmail);

            while (true) ;
            
        }

        private static void outlook_newmail(string s)
        {
            string subject = "leer";
            Outlook.MailItem mailItem = null;

            Console.WriteLine("Entered Event with inboxitems: "+ inboxFolder.Items.Count);

            for (int i = (inboxFolder.Items.Count); i > 0; i--)
            {
                mailItem = inboxFolder.Items[i] as Outlook.MailItem;
               
                
                if (mailItem.Subject.StartsWith("DevDB 1", System.StringComparison.CurrentCultureIgnoreCase))
                {
                    subject = mailItem.Subject;
                    einbuchung_vorgang(subject);
                    mailItem.Delete();
                }
                else if (mailItem.Subject.StartsWith("DevDB 2", System.StringComparison.CurrentCultureIgnoreCase))
                {
                    subject = mailItem.Subject;
                    ausbuchung_vorgang(subject);
                    mailItem.Delete();
                }
                else
                {
                    mailItem.Delete();
                    Console.WriteLine("Subject didn't match with rules!");
                }
            }
            Console.WriteLine("Left Event");
        }

        private static void einbuchung_vorgang(string s)
        {
            Console.WriteLine("Einbuchung: " + s + " erfolgreich erledigt");
        }

        private static void ausbuchung_vorgang(string s)
        {
            Console.WriteLine("Einbuchung: " + s + " erfolgreich erledigt");
        }
       
       
    }
}
