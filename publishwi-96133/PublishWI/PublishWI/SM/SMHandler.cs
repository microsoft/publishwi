using System;
using System.Collections.Generic;
using Microsoft.Win32;

using Microsoft.EnterpriseManagement;
using Microsoft.EnterpriseManagement.Common;
using Microsoft.EnterpriseManagement.Configuration;

namespace PublishWI
{
    public class SMHandler
    {

        public const string SystemWorkitemLibraryMP = "System.WorkItem.Library";
        public const string WIType = "System.WorkItem";

        private EnterpriseManagementGroup mg;
        private ManagementPackClass wiClass;

        public SMHandler()
        {

            String strSMServerName = null;
            try
            {
                strSMServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
            }
            catch
            {
                Console.WriteLine("Cannot read the name of the SM Server from Registry.");
                Environment.Exit(2);
            }

            try
            {
                mg = new EnterpriseManagementGroup(strSMServerName);
            }
            catch
            {
                Console.WriteLine("Cannot connect to SM Server {0}", strSMServerName);
                Environment.Exit(2);
            }

            ManagementPack systemWorkitemMP = GetManagementPackByName("System.WorkItem.Library", this.mg);

            wiClass = systemWorkitemMP.GetClass(WIType);

        }

        public static ManagementPack GetManagementPackByName(string strName, EnterpriseManagementGroup emg)
        {
            ManagementPackCriteria mpc = new ManagementPackCriteria(String.Format("Name = '{0}'", strName));
            IList<ManagementPack> ListManagementPacks = emg.ManagementPacks.GetManagementPacks(mpc);
            if (ListManagementPacks.Count > 0)
            {
                return (ListManagementPacks[0]);
            }
            else
            {
                return null;
            }
        }

        public WorkItem DownloadWIfromSM(string Id)
        {

            EnterpriseManagementObject wiInSM = null;
            WorkItem wi = new WorkItem();

            EnterpriseManagementObjectCriteria criteria = new EnterpriseManagementObjectCriteria(String.Format("Name='{0}'", Id), wiClass);

            IObjectReader<EnterpriseManagementObject> reader = mg.EntityObjects.
                GetObjectReader<EnterpriseManagementObject>(criteria, ObjectQueryOptions.Default);

            if (reader.Count == 1)
            {    // Found WI
                wiInSM = reader.GetRange(0, 1)[0];
            }
            else
            {
                Console.WriteLine("Can't find Work Item {0}", Id);
                Environment.Exit(2);
            }

            wi.Id = Convert.ToString(wiInSM[wiClass, "Id"].Value);
            wi.Title = Convert.ToString(wiInSM[wiClass, "Title"].Value);
            wi.Description = Convert.ToString(wiInSM[wiClass, "Description"].Value);
            wi.ContactMethod = Convert.ToString(wiInSM[wiClass, "ContactMethod"].Value);
            wi.DisplayName = Convert.ToString(wiInSM[wiClass, "DisplayName"].Value);
            wi.ScheduledStartDate = Convert.ToDateTime(wiInSM[wiClass, "ScheduledStartDate"].Value).ToLocalTime();
            wi.ScheduledEndDate = Convert.ToDateTime(wiInSM[wiClass, "ScheduledEndDate"].Value).ToLocalTime();
            wi.ActualStartDate = Convert.ToDateTime(wiInSM[wiClass, "ActualStartDate"].Value).ToLocalTime();
            wi.ActualEndDate = Convert.ToDateTime(wiInSM[wiClass, "ActualEndDate"].Value).ToLocalTime();
            //            wi.LastModified     = Convert.ToDateTime(wiInSM[null, "LastModified"].Value).ToLocalTime();  // doesnt' work - this one is from System.Entity

            return wi;
        }
    }
}