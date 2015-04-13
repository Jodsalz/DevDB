using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using NetOffice.OutlookApi.Enums;
using System.Text.RegularExpressions;

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
        static MySqlConnection devdb;

        static void Main(string[] args)
        {

            // start outlook
            outlookApplication = new Outlook.Application();
            devdb = new MySqlConnection(@"Server=127.0.0.1;Uid=root;Pwd=;Database=devdb;");
            
            // get inbox
            outlookNS = outlookApplication.Session;
            inboxFolder = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["Inbox"];

            einbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Einbuchung"];
            Outlook.MAPIFolder ausbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Ausbuchung"];


            outlookApplication.NewMailExEvent += new Outlook.Application_NewMailExEventHandler(outlook_newmail);                // Initializing both Event Handlers
            outlookApplication.Session.SyncObjects[1].SyncEndEvent += new Outlook.SyncObject_SyncEndEventHandler(sync_end);

            while (true)
            {
                
               
            }
            
        }

        private static void outlook_newmail(string s)           // Function starts through event (NewMailExEvent)
        {
            Console.WriteLine("enter newmail event");
            Console.WriteLine("starting sendandreceiv");
            outlookNS.SendAndReceive(false);
        }

        private static void sync_end()                          // Function starts through event (SyncEndEvent)
        {
            Console.WriteLine("enter syncend event");

            string subject = "leer";
            string body = "leer";
            Outlook.MailItem mailItem = null;

            Console.WriteLine("Entered sync_end Event with inboxitems: " + inboxFolder.Items.Count);

            for (int i = inboxFolder.Items.Count; i > 0; i--)
            {
                mailItem = inboxFolder.Items[i] as Outlook.MailItem;


                // Downloads Body if not downloaded yet (IMAP Problems)
                if (mailItem.DownloadState != OlDownloadState.olFullItem)
                {
                    Console.WriteLine("again");
                    mailItem.MarkForDownload = OlRemoteStatus.olMarkedForDownload;
                    outlookNS.SendAndReceive(false);
                    break;                                                                      // Replace with threads maybe?
                }

                // Calls Database for Insert
                if (mailItem != null && mailItem.Subject.StartsWith("DevDB Einbuchung", System.StringComparison.CurrentCultureIgnoreCase) && mailItem.DownloadState == OlDownloadState.olFullItem)
                {
                    subject = mailItem.Subject;
                    body = mailItem.Body;
                    einbuchung_vorgang(body);
                    mailItem.Delete();
                }

                // Calls Database for Delete
                else if (mailItem != null && mailItem.Subject.StartsWith("DevDB 2", System.StringComparison.CurrentCultureIgnoreCase) && mailItem.DownloadState == OlDownloadState.olFullItem)
                {
                    subject = mailItem.Subject;
                    body = mailItem.Body;
                    ausbuchung_vorgang(body);
                    mailItem.Delete();
                }
                else
                {
                    Console.WriteLine("doesn't match rules... deleted...");
                    mailItem.Delete();
                }

            }
            Console.WriteLine("Left Event");
        }

        private static void passwort(string s)
        {
            Console.WriteLine("Einbuchung: ...");

            string stm = "SELECT kennwort FROM mitarbeiter WHERE name = '" + s.Replace("\r\n", "") + "'";
            try
            {
                devdb.Open();
            }
            catch (MySqlException ex)
            {
                Console.WriteLine("error opening connection {0}", ex.ToString());
            }
            MySqlCommand cmd = new MySqlCommand(stm, devdb);
            string version = Convert.ToString(cmd.ExecuteScalar());
            Console.WriteLine(version);
            if (devdb != null)
            {
                devdb.Close();
            }
        }

        private static void einbuchung_vorgang(string s)
        {
            string von_datum, bis_datum, device_id, ausgeliehen_an, ausgeliehen_an_tel, ausgeliehen_an_email, ausgeliehen_von, kommentar;
            string[] lines = Regex.Split(s, "\r\n");
            von_datum = lines[0].Substring(lines[0].IndexOf(":") + 2); // Lese nur rechten Teil nach : aus, IndexOf liefert position von : -> +2 wegen leerzeichen 
            bis_datum = lines[1].Substring(lines[1].IndexOf(":") + 2);
            device_id = lines[2].Substring(lines[2].IndexOf(":") + 2);
            ausgeliehen_an = lines[3].Substring(lines[3].IndexOf(":") + 2);
            ausgeliehen_an_tel = lines[4].Substring(lines[4].IndexOf(":") + 2);
            ausgeliehen_an_email = lines[5].Substring(lines[5].IndexOf(":") + 2);
            ausgeliehen_von = lines[6].Substring(lines[6].IndexOf(":") + 2);
            kommentar = lines[7].Substring(lines[7].IndexOf(":") + 2);
            Console.WriteLine(von_datum + "\n" + bis_datum + "\n" + device_id + "\n" + ausgeliehen_an + "\n" + ausgeliehen_an_tel + "\n" + ausgeliehen_an_email + "\n" + ausgeliehen_von + "\n" + kommentar);
        }

        private static void ausbuchung_vorgang(string s)
        {
            Console.WriteLine("Einbuchung: " + s + ": " + s + " erfolgreich erledigt");
        }
       
       
    }
}
