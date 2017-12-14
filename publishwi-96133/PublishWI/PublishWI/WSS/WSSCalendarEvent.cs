using System;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;

namespace PublishWI
{
    public class WSSCalendarEvent
    {

        public int Id { get; set; }
        public Guid GUID { get; set; }
        public string Title { get; set; }
        public string Location { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string Description { get; set; }
        public Enum Category { get; set; }
        public DateTime Modified { get; set; }
        public int EventType { get; set; }

        public void ReadFromItem(ListItem item)
        {
            Id = item.Id;
            GUID = new Guid(item.FieldValuesAsText["GUID"]);
            Title = item.DisplayName;
            Description = item.FieldValuesAsText["Description"];
            Location = item.FieldValuesAsText["Location"];
            StartTime = Convert.ToDateTime(item.FieldValuesAsText["EventDate"]);
            EndTime = Convert.ToDateTime(item.FieldValuesAsText["EndDate"]);
            Modified = Convert.ToDateTime(item.FieldValuesAsText["Modified"]);
            EventType = Convert.ToInt32(item.FieldValuesAsText["EventType"]);
            // public Enum Category { get; set; }

        }

        public void PrintEvent()
        {
            Console.WriteLine("ID: {0}", Id);
            Console.WriteLine("GUID: {0}", GUID);
            Console.WriteLine("Title: {0}", Title);
            Console.WriteLine("Description: {0}", Description);
            Console.WriteLine("Location: {0}", Location);
            Console.WriteLine("Start Date: {0}", StartTime.ToString());
            Console.WriteLine("End Date: {0}", EndTime.ToString());
            Console.WriteLine("Modified: {0}", Modified.ToString());
            Console.WriteLine("Event Type: {0}", EventType);

        }
    }
}
