using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PublishWI
{
    public class WorkItem
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public DateTime ActualStartDate { get; set; }
        public DateTime ActualEndDate { get; set; }
        public string ContactMethod { get; set; }
        public DateTime ScheduledStartDate { get; set; }
        public DateTime ScheduledEndDate { get; set; }
        public string DisplayName { get; set; }
        public DateTime LastModified { get; set; }


        public void Print()
        {
            Console.WriteLine("Id: {0}", Id);
            Console.WriteLine("Title: {0}", Title);
            Console.WriteLine("Description: {0}", Description);
            Console.WriteLine("Contact Method: {0}", ContactMethod);
            Console.WriteLine("Scheduled Start: {0}", ScheduledStartDate.ToString());
            Console.WriteLine("Scheduled End: {0}", ScheduledEndDate.ToString());
        }
    }
}
