using System;

using Microsoft.EnterpriseManagement;
using Microsoft.EnterpriseManagement.Common;
using Microsoft.EnterpriseManagement.Configuration;
using Microsoft.EnterpriseManagement.Packaging;


namespace PublishWI
{
    class Program
    {
        static void Main(string[] args)
        {

            string wiID="";
            string wssServer="";
            string wssList="";

            if (args.Length == 3) { 
                wiID = args[0];
                wssServer = args[1];
                wssList = args[2];
            } else {
                const string HelpAbout =
                    "Usage: PublishWI.exe WIID URL CalendarName\n" +
                    "          WIID: ID of the Work Item to publish (for example, CR3333)\n" +
                    "           URL: URL of the SharePoint site, such as http://www.sharepoint.com \n" +
                    "  CalendarName: The name of the sharepoint calendar to publish WI to. \n";

                Console.Write(HelpAbout);
                Environment.Exit(1);
            }

            WorkItem workItem;
            SMHandler serviceManager = new SMHandler();
            workItem = serviceManager.DownloadWIfromSM(wiID);

            WSSHandler calendar = new WSSHandler(wssServer, wssList);

            if (calendar.SearchForEventByTitleOnWSS(workItem.DisplayName))
            {
                calendar.UpdateExistingEventOnWSS(workItem);
                Console.WriteLine("Event for Work Item {0} successfully updated\n", wiID);
            }
            else
            {
                calendar.CreateNewEventOnWSS(workItem);
                Console.WriteLine("Event for Work Item {0} successfully created\n", wiID);
            }

            workItem.Print();
        }
    }
}

