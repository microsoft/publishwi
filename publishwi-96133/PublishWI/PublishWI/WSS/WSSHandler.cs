using System;

using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;


namespace PublishWI
{
    public class WSSHandler
    {
        private ClientContext clientContext;
        private SP.List oList;
        private ListItem oListItem;
        private WSSCalendarEvent evt;

        public WSSHandler(string sURL, string sListName)
        {
            try
            {
                clientContext = new ClientContext(sURL);
                oList = clientContext.Web.Lists.GetByTitle(sListName);
                evt = new WSSCalendarEvent();
            }
            catch
            {
                Console.WriteLine("Can't connect to the {0} list on {1}", sListName, sURL);
                Environment.Exit(3);
            }
        }

        public bool SearchForEventByTitleOnWSS(string strTitle)
        {
            
            CamlQuery camlQuery = new CamlQuery();
                
            camlQuery.ViewXml = String.Format(
                "<View><Query><Where><Eq><FieldRef Name='Title'/>" +
                "<Value Type='Text'>{0}</Value></Eq></Where></Query>" + 
                "</View>",
                strTitle);
            
            ListItemCollection collListItem = oList.GetItems(camlQuery);
            clientContext.Load(collListItem);  
            clientContext.ExecuteQuery();

            if (collListItem.Count > 1)
            {
                Console.WriteLine("Error: The list contains {0} calendar events with the same title \"{1}\"", collListItem.Count, strTitle);
                Environment.Exit(3);
            }
            else if (collListItem.Count == 0)
                return false;

            oListItem = collListItem[0];
            return true;
        }
        
        public void CreateNewEventOnWSS(WorkItem wi)
        {

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            oListItem = oList.AddItem(itemCreateInfo);
            oListItem["Title"] = wi.DisplayName;
            oListItem["Description"] = wi.Description;
            oListItem["EventDate"] = wi.ScheduledStartDate;
            oListItem["EndDate"] = wi.ScheduledEndDate;
            oListItem.Update();

            try
            {
                clientContext.ExecuteQuery();
            }
            catch
            {
                Console.WriteLine("Can't create event on the sharepoint calendar");
                Environment.Exit(3);
            }
 
        
        }

        public void UpdateExistingEventOnWSS(WorkItem wi)
        {
            oListItem["Title"] = wi.DisplayName;
            oListItem["Description"] = wi.Description;
            oListItem["EventDate"] = wi.ScheduledStartDate;
            oListItem["EndDate"] = wi.ScheduledEndDate;
            oListItem.Update();

            try
            {
                clientContext.ExecuteQuery();
            }
            catch
            {
                Console.WriteLine("Can't update event on the sharepoint calendar");
                Environment.Exit(3);
            }

        }        

    } 
}

