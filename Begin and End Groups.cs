using System;
using System.Collections.Generic;
using System.Text;

namespace Custom_Report_Debugger_2._0
{
    class BeginAndEndGroups
    {
        public static (List<string>, List<string>) CheckAllBeginGroupsHaveEndGroups(List<string> mergefieldsToTest)
        {
            var failedMergefields_MissingEndGroup = new List<string>();
            var failedMergefields_MissingBeginGroup = new List<string>();
            var beginGroups = new List<string>();
            var endGroups = new List<string>();
            var standardMergefields = new List<string>();
            var openBeginGroups = new List<string>();

            foreach (var field in mergefieldsToTest)
            {
                if (field.Contains("BeginGroup"))
                {
                    var splitBegin = field.Split(':');
                    beginGroups.Add(splitBegin[1]);
                }
                else if (field.Contains("EndGroup"))
                {
                    var splitEnd = field.Split(':');
                    endGroups.Add(splitEnd[1]);
                }
                else
                {
                    standardMergefields.Add(field);
                }
            }
            foreach (var field in beginGroups)
            {
                openBeginGroups.Add(field);
                if (endGroups.Contains(field))
                {
                    openBeginGroups.Remove(field);
                    endGroups.Remove(field);
                }
            }
            foreach (var field in openBeginGroups)
            {
                failedMergefields_MissingEndGroup.Add("{MERGEFIELD BeginGroup:" + field);
            }
            foreach (var field in endGroups)
            {
                failedMergefields_MissingBeginGroup.Add("{MERGEFIELD EndGroup:" + field);
            }
            return (failedMergefields_MissingBeginGroup, failedMergefields_MissingEndGroup);
        }
    }
}