using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Globalization;

namespace Portfolio.Tools.Reporter
{
    /// <summary>
    /// System utility functions for report variable substitution.
    /// Demonstrates: Active Directory queries, recursive group membership, LDAP.
    /// </summary>
    static class SystemLibrary
    {
        public static string CurrentDatetime()
        {
            return DateTime.Now.ToString("M/d/yyyy h:mm:ss tt");
        }

        public static string FirstOfMonth(string MonthsAgo)
        {
            string returnValue = "";

            int.TryParse(MonthsAgo, out int monthsAgo);
            returnValue = DateTime.Now.AddDays(1 - DateTime.Now.Day).AddMonths(-monthsAgo).Date.ToString();

            return returnValue;
        }

        public static bool IsMatch(string String1, string String2)
        {
            bool returnValue = false;

            returnValue = CultureInfo.CurrentCulture.CompareInfo.Compare(String1, String2, CompareOptions.IgnoreCase) == 0;

            return returnValue;
        }

        /// <summary>
        /// Get all members of AD groups, including nested group members.
        /// </summary>
        public static string ADGroupMembers(string ADGroups)
        {
            string returnValue = "";

            returnValue = GetADGroupMembers(ADGroups, true);

            if (returnValue == "")
            {
                throw new Exception($"No Active Directory Members found for security group(s): {ADGroups}");
            }

            return returnValue;
        }

        /// <summary>
        /// Recursively enumerate AD group members using LDAP.
        /// Demonstrates: DirectorySearcher, LDAP filters, recursive algorithms.
        /// </summary>
        private static string GetADGroupMembers(string adGroups, bool recursive = true, string memberList = "")
        {
            DirectoryEntry directoryEntry;
            PropertyValueCollection members;
            DirectoryEntry member;
            string memberValue;

            adGroups = adGroups.Trim().Trim(',') + ",";
            string[] ADGroupsArray = adGroups.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            if (ADGroupsArray.Length > 0)
            {
                // Build LDAP group filter
                string searchFilter = "(&(objectClass=Group)";
                searchFilter += "(|";

                foreach (var group in ADGroupsArray)
                {
                    searchFilter += string.Format("(SAMAccountName={0})", group);
                }

                searchFilter += "))";

                // Execute filtered search
                var searcher = new DirectorySearcher();
                searcher.Filter = searchFilter;
                searcher.PropertiesToLoad.Add("member");
                var searchResults = searcher.FindAll();

                // Add members of each group to collection
                foreach (SearchResult result in searchResults)
                {
                    directoryEntry = (DirectoryEntry)result.GetDirectoryEntry();
                    members = directoryEntry.Properties["member"];

                    foreach (string name in members)
                    {
                        member = new DirectoryEntry(string.Format("LDAP://{0}", name));
                        memberValue = member.Properties["sAMAccountName"].Value.ToString();

                        if (member.SchemaClassName == "user")
                            memberList += $"{(string.IsNullOrEmpty(memberList) ? "" : ",")}{memberValue}";

                        if (recursive && member.SchemaClassName == "group")
                            memberList = GetADGroupMembers(memberValue, recursive, memberList);
                    }
                }
            }

            return memberList;
        }
    }
}
