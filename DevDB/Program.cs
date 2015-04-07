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
            Outlook.Application outlookApplication = new Outlook.Application();


            // get inbox
            Outlook._NameSpace outlookNS = outlookApplication.Session;
            Outlook.MAPIFolder inboxFolder = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["Posteingang"];
            /*
            Outlook.MAPIFolder einbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Einbuchung"];
            Outlook.MAPIFolder ausbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Ausbuchung"];
            */
            Outlook.MAPIFolder einbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["Posteingang"].Folders["DevDB"].Folders["Einbuchung"];
            Outlook.MAPIFolder ausbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["Posteingang"].Folders["DevDB"].Folders["Ausbuchung"];

            foreach (COMObject item in inboxFolder.Items)
            {
                Outlook.MailItem mailItem = item as Outlook.MailItem;
                if (mailItem != null && mailItem.Subject == "DevDB Einbuchung")
                {
                    body = mailItem.Body;
                    mailItem.Move(einbuchung);
                    mailItem.Save();
                }
                else if (mailItem != null && mailItem.Subject == "DevDB Ausbuchung")
                {
                    body = mailItem.Body;
                    mailItem.Move(ausbuchung);
                    mailItem.Save();
                }
            }

            Console.WriteLine(body);

            outlookApplication.Dispose();
            Console.ReadKey();
        }
    }
}
