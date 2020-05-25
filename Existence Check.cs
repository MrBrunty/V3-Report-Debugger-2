using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace Custom_Report_Debugger_2._0
{
    public class ExistenceCheck
    {
        public static List<string> CheckExistenceOfMergefield(List<string> mergefieldsToTest, List<string> legalMergefields, string logFile)
        {
            var failedMergefields_doNotExist = new List<string>();

            File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine + "The following fields were checked for existence in the Mergefiled Database:" + Environment.NewLine);
            foreach (var field in mergefieldsToTest)
            {
                if (!legalMergefields.Contains(field))
                {

                    failedMergefields_doNotExist.Add(field);
                    File.AppendAllText(logFile, Environment.NewLine + field + " Could not be found within the Mergefield database");
                }
                else
                {
                    File.AppendAllText(logFile, Environment.NewLine + field + " Was found");
                }
            }

            return failedMergefields_doNotExist;
        }

        public static (List<string>, List<string>) DefineLegalMergefields(string[] configMergefields) //This is a little pointless, however it converts the fields to a list format and stops us from having to iterate to duplicate mergefields later on.
        {
            var listOfFields_NoWhiteSpace = new List<string>();
            var uniqueMergefields = new List<string>();
            var duplicateMergefields = new List<string>();


            foreach (var field in configMergefields)
            {
                listOfFields_NoWhiteSpace.Add(field.Replace("\n", "").Trim());
            }

            foreach (var field in listOfFields_NoWhiteSpace)
            {
                if (uniqueMergefields.Contains(field))
                {
                    if (duplicateMergefields.Contains(field))
                    {
                        continue;
                    }
                    else
                    {
                        duplicateMergefields.Add(field);
                    }
                }
                else
                {
                    uniqueMergefields.Add(field);
                }
            }
            return (uniqueMergefields, duplicateMergefields);
        }

        public static List<string> RemoveAdditionalSpacingAndUnneccearyFields(List<string> mergefieldsToTest, string logFile)
        {
            var amendedMergefields = new List<string>();
            var usableMergefields = new List<string>();
            bool usable;
            var ignoreList = HierarchyChecker.ConfigImport(ConfigurationManager.AppSettings.Get("SpecificTextToIgnore").Split(','));

            foreach (var field in mergefieldsToTest)
            {
                var amended = field.Replace("{  ", "{").Replace("{ ", "{").Replace("  }", "}").Replace(" }", "}").Replace("MERGEFIELD  ", "MERGEFIELD ").Replace("  \\* MERGEFORMAT", "").Replace(" \\* MERGEFORMAT", "").Replace("\\* MERGEFORMAT", "");
                amendedMergefields.Add(amended);
            }

            foreach (var field in amendedMergefields)
            {
                usable = true;
                foreach (var parameter in ignoreList)
                {

                    if (field.Contains(parameter) | field.Contains("\""))
                    {
                        File.AppendAllText(logFile, Environment.NewLine + field + " Was removed from the test data due to \"" + parameter + "\" appearing in it");
                        usable = false;
                    }
                }
                if (!field.Contains("}") && !field.Contains("{"))
                {
                    usable = false;
                }
                if (usable == true)
                {
                    usableMergefields.Add(field);
                }
            }

            return usableMergefields;
        }

    }
}