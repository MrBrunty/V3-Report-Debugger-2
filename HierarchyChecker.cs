using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace Custom_Report_Debugger_2._0
{
    class HierarchyChecker
    {
        public static List<string> ConfigImport(string[] config)
        {
            var group = new List<string>();

            foreach (var field in config)
            {
                group.Add(field.Trim().Replace("\n", ""));
            }
            return group;
        }

        public static List<string> HierarchyManagement(List<string> mergefieldsToTest, string logFile)
        {            
            //Load all configs
            var report_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Report").Split(','));
            var managedAssetSummary_Group = ConfigImport(ConfigurationManager.AppSettings.Get("ManagedAssetSummary").Split(','));
            var managedAssetSummary_Plans_Group = ConfigImport(ConfigurationManager.AppSettings.Get("ManagedAssetSummary_Plans").Split(','));
            var managedAssetSummary_AllPlans_Group = ConfigImport(ConfigurationManager.AppSettings.Get("ManagedAssetSummary_AllPlans").Split(','));
            var managedAssetSummary_SubPlans_Group = ConfigImport(ConfigurationManager.AppSettings.Get("ManagedAssetSummary_SubPlans").Split(','));
            var managedAssetSummary_CostsAndChargesSummary_Group = ConfigImport(ConfigurationManager.AppSettings.Get("ManagedAssetSummary_CostsAndChargesSummary").Split(','));
            var managedAssetSummary_CostsAndCharges_Group = ConfigImport(ConfigurationManager.AppSettings.Get("ManagedAssetSummary_CostsAndCharges").Split(','));

            var protections_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections").Split(','));
            var protections_Plans_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_Plans").Split(','));
            var protections_FundsWithoutPortfolio_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_FundsWithoutPortfolio").Split(','));
            var protections_AssetClasses_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_AssetClasses").Split(','));
            var protections_Regions_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_Regions").Split(','));
            var protections_Covers_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_Covers").Split(','));
            var protections_Funds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_Funds").Split(','));
            var protections_Detailed_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_Detailed").Split(','));
            var protections_Summarised_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_Summarised").Split(','));
            var protections_CoverDetails_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_CoverDetails").Split(','));
            var protections_LookThrough_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_LookThrough").Split(','));
            var protections_Assets_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Protections_Assets").Split(','));

            var mortgages_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Mortgages").Split(','));
            var mortgages_Plans_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Mortgages_Plans").Split(','));
            var mortgages_FundsWithoutPortfolio_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Mortgages_FundsWithoutPortfolio").Split(','));
            var mortgages_LookThrough_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Mortgages_LookThrough").Split(','));
            var mortgages_Assets_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Mortgages_Assets").Split(','));
            var mortgages_Funds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Mortgages_Funds").Split(','));

            var generalInsurances_Group = ConfigImport(ConfigurationManager.AppSettings.Get("GeneralInsurances").Split(','));
            var generalInsurances_Plans_Group = ConfigImport(ConfigurationManager.AppSettings.Get("GeneralInsurances_Plans").Split(','));
            var generalInsurances_FundsWithoutPortfolio_Group = ConfigImport(ConfigurationManager.AppSettings.Get("GeneralInsurances_FundsWithoutPortfolio").Split(','));
            var generalInsurances_LookThrough_Group = ConfigImport(ConfigurationManager.AppSettings.Get("GeneralInsurances_LookThrough").Split(','));
            var generalInsurances_Funds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("GeneralInsurances_Funds").Split(','));
            var generalInsurances_Assets_Group = ConfigImport(ConfigurationManager.AppSettings.Get("GeneralInsurances_Assets").Split(','));

            var fundXRay_Group = ConfigImport(ConfigurationManager.AppSettings.Get("FundXRay").Split(','));
            var fundXRay_AssetClasses_Group = ConfigImport(ConfigurationManager.AppSettings.Get("FundXRay_AssetClasses").Split(','));
            var fundXRay_Regions_Group = ConfigImport(ConfigurationManager.AppSettings.Get("FundXRay_Regions").Split(','));
            var fundXRay_Summarised_Group = ConfigImport(ConfigurationManager.AppSettings.Get("FundXRay_Summarised").Split(','));
            var fundXRay_Detailed_Group = ConfigImport(ConfigurationManager.AppSettings.Get("FundXRay_Detailed").Split(','));

            var taxWrappers_Group = ConfigImport(ConfigurationManager.AppSettings.Get("TaxWrappers").Split(','));
            var taxWrappers_Plans_Group = ConfigImport(ConfigurationManager.AppSettings.Get("TaxWrappers_Plans").Split(','));

            var investments_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments").Split(','));
            var investments_PlansWithFunds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_PlansWithFunds").Split(','));
            var investments_FundsWithoutPortfolio_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_FundsWithoutPortfolio").Split(','));
            var investments_Funds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_Funds").Split(','));
            var investments_LookThrough_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_LookThrough").Split(','));
            var investments_Assets_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_Assets").Split(','));
            var investments_FundsWithPortfolio_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_FundsWithPortfolio").Split(','));
            var investments_Portfolios_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_Portfolios").Split(','));
            var investments_SubPlansWithFunds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_SubPlansWithFunds").Split(','));
            var investments_CostsAndChargesSummary_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_CostsAndChargesSummary").Split(','));
            var investments_CostsAndCharges_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_CostsAndCharges").Split(','));
            var investments_AssetClasses_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_AssetClasses").Split(','));
            var investments_Summarised_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_Summarised").Split(','));
            var investments_Detailed_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_Detailed").Split(','));
            var investments_Regions_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_Regions").Split(','));
            var investments_SubPlansWithoutFunds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_SubPlansWithoutFunds").Split(','));
            var investments_AllPlans_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_AllPlans").Split(','));
            var investments_PlansWithoutFunds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Investments_PlansWithoutFunds").Split(','));

            var pensions_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions").Split(','));
            var pensions_PlansWithFunds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_PlansWithFunds").Split(','));
            var pensions_FundsWithoutPortfolio_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_FundsWithoutPortfolio").Split(','));
            var pensions_Funds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_Funds").Split(','));
            var pensions_LookThrough_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_LookThrough").Split(','));
            var pensions_Assets_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_Assets").Split(','));
            var pensions_FundsWithPortfolio_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_FundsWithPortfolio").Split(','));
            var pensions_Portfolios_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_Portfolios").Split(','));
            var pensions_SubPlansWithFunds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_SubPlansWithFunds").Split(','));
            var pensions_CostsAndChargesSummary_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_CostsAndChargesSummary").Split(','));
            var pensions_CostsAndCharges_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_CostsAndCharges").Split(','));
            var pensions_AssetClasses_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_AssetClasses").Split(','));
            var pensions_Summarised_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_Summarised").Split(','));
            var pensions_Detailed_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_Detailed").Split(','));
            var pensions_Regions_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_Regions").Split(','));
            var pensions_SubPlansWithoutFunds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_SubPlansWithoutFunds").Split(','));
            var pensions_AllPlans_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_AllPlans").Split(','));
            var pensions_PlansWithoutFunds_Group = ConfigImport(ConfigurationManager.AppSettings.Get("Pensions_PlansWithoutFunds").Split(','));

            int indexOfLastOpenGroup;
            string currentGroup = "Null";
            var activeGroups = new List<string>();
            var failedMergefields_Hierarchy = new List<string>();
            File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine);

            foreach (var field in mergefieldsToTest)
            {
                if (field.Contains("BeginGroup") | field.Contains("EndGroup"))
                {
                    if (currentGroup == "{MERGEFIELD BeginGroup:Report}")
                    {
                        if (!field.Contains("{MERGEFIELD EndGroup:Report}") && !field.Contains("Investments") && !field.Contains("ManagedAssetSummary") && !field.Contains("Protections") && !field.Contains("Mortgages") && !field.Contains("GeneralInsurances") && !field.Contains("Pensions") && !field.Contains("TaxWrappers") && !field.Contains("FundXRay"))
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:ManagedAssetSummary}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Plans}" && field != "{MERGEFIELD BeginGroup:AllPlans}" && field != "{MERGEFIELD EndGroup:ManagedAssetSummary}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Plans}")
                    {
                        if (field != "{MERGEFIELD EndGroup:Plans}" && field != "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}" && field != "{MERGEFIELD BeginGroup:SubPlans}" && field != "{MERGEFIELD BeginGroup:CostsAndChargesSummary}" && field != "{MERGEFIELD BeginGroup:AssetClasses}" && field != "{MERGEFIELD BeginGroup:Regions}" && field != "{MERGEFIELD BeginGroup:Covers}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:CostsAndChargesSummary}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:CostsAndCharges}" && field != "{MERGEFIELD EndGroup:CostsAndChargesSummary}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:CostsAndCharges}")
                    {
                        if (field != "{MERGEFIELD EndGroup:CostsAndChargesSummary}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:SubPlans}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:CostsAndChargesSummary}" && field != "{MERGEFIELD EndGroup:SubPlans}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:AllPlans}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:CostsAndChargesSummary}" && field != "{MERGEFIELD BeginGroup:Regions}" && field != "{MERGEFIELD BeginGroup:AssetClasses}" && field != "{MERGEFIELD EndGroup:AllPlans}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Investments}")
                    {

                        if (field != "{MERGEFIELD BeginGroup:PlansWithFunds}" && field != "{MERGEFIELD BeginGroup:PlansWithoutFunds}" && field != "{MERGEFIELD BeginGroup:AllPlans}" && field != "{MERGEFIELD EndGroup:Investments}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }

                    else if (currentGroup == "{MERGEFIELD BeginGroup:PlansWithFunds}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}" && field != "{MERGEFIELD BeginGroup:FundsWithPortfolio}" && field != "{MERGEFIELD BeginGroup:SubPlansWithFunds}" && field != "{MERGEFIELD BeginGroup:SubPlansWithoutFunds}" && field != "{MERGEFIELD BeginGroup:AssetClasses}" && field != "{MERGEFIELD BeginGroup:Regions}" && field != "{MERGEFIELD BeginGroup:CostsAmdChargesSummary}" && field != "{MERGEFIELD EndGroup:PlansWithFunds}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Funds}" && field != "{MERGEFIELD EndGroup:FundsWithoutPortfolio}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Funds}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:LookThrough}" && field != "{MERGEFIELD EndGroup:Funds}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:LookThrough}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Assets}" && field != "{MERGEFIELD EndGroup:LookThrough}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Assets}")
                    {
                        if (field != "{MERGEFIELD EndGroup:Assets}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundsWithPortfolio}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Portfolios}" && field != "{MERGEFIELD EndGroup:FundsWithPortfolio}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Portfolios}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Funds}" && field != "{MERGEFIELD EndGroup:Portfolios}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:SubPlansWithFunds}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}" && field != "{MERGEFIELD BeginGroup:FundsWithPortfolio}" && field != "{MERGEFIELD BeginGroup:AssetClasses}" && field != "{MERGEFIELD BeginGroup:Regions}" && field != "{MERGEFIELD BeginGroup:CostsAndChargesSummary}" && field != "{MERGEFIELD EndGroup:SubPlansWithFunds}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:AssetClasses}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Summarised}" && field != "{MERGEFIELD BeginGroup:Detailed}" && field != "{MERGEFIELD EndGroup:AssetClasses}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Summarised}")
                    {
                        if (field != "{MERGEFIELD EndGroup:Summarised}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Detailed}")
                    {
                        if (field != "{MERGEFIELD EndGroup:Detailed}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Regions}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Summarised}" && field != "{MERGEFIELD EndGroup:Regions}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:SubPlansWithoutFunds}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}" && field != "{MERGEFIELD BeginGroup:FundsWithPortfolio}" && field != "{MERGEFIELD BeginGroup:AssetClasses}" && field != "{MERGEFIELD BeginGroup:Regions}" && field != "{MERGEFIELD BeginGroup:CostsAndChargesSummary}" && field != "{MERGEFIELD EndGroup:SubPlansWithoutFunds}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:PlansWithoutFunds}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}" && field != "{MERGEFIELD BeginGroup:FundsWithPortfolio}" && field != "{MERGEFIELD BeginGroup:SubPlansWithFunds}" && field != "{MERGEFIELD BeginGroup:SubPlansWithoutFunds}" && field != "{MERGEFIELD BeginGroup:AssetClasses}" && field != "{MERGEFIELD BeginGroup:Regions}" && field != "{MERGEFIELD BeginGroup:CostsAmdChargesSummary}" && field != "{MERGEFIELD EndGroup:PlansWithoutFunds}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Pensions}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:PlansWithFunds}" && field != "{MERGEFIELD BeginGroup:PlansWithoutFunds}" && field != "{MERGEFIELD BeginGroup:AllPlans}" && field != "{MERGEFIELD EndGroup:Pensions}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Mortgages}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Plans}" && field != "{MERGEFIELD EndGroup:Mortgages}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Protections}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Plans}" && field != "{MERGEFIELD EndGroup:Protections}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:GeneralInsurances}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Plans}" && field != "{MERGEFIELD EndGroup:GeneralInsurances}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:TaxWrappers}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:Plans}" && field != "{MERGEFIELD BeginGroup:PlansWithoutFunds}" && field != "{MERGEFIELD EndGroup:TaxWrappers}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }

                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundXRay}")
                    {
                        if (field != "{MERGEFIELD BeginGroup:AssetClasses}" && field != "{MERGEFIELD BeginGroup:Regions}" && field != "{MERGEFIELD EndGroup:FundXRay}")
                        {
                            failedMergefields_Hierarchy.Add(field);
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                        }
                    }


                    if (field.Contains("BeginGroup"))
                    {
                        var splitBegin = field.Split(':');
                        activeGroups.Add(splitBegin[1]);
                    }
                    else if (field.Contains("EndGroup"))
                    {
                        var splitEnd = field.Split(':');
                        activeGroups.Remove(splitEnd[1]);
                    }


                    indexOfLastOpenGroup = activeGroups.Count - 1;
                    if (indexOfLastOpenGroup < 0)
                    {
                        currentGroup = "Report}";
                    }
                    else
                    {
                        currentGroup = activeGroups[indexOfLastOpenGroup];
                    }
                    currentGroup = "{MERGEFIELD BeginGroup:" + currentGroup;
                }

                else
                {
                    foreach (var group in activeGroups)
                    {
                        var neatGroup = group.Replace("}", "");
                        File.AppendAllText(logFile, neatGroup + " --> ");
                    }


                    //REPORT
                    if (currentGroup == "{MERGEFIELD BeginGroup:Report}")
                    {
                        if (!report_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }

                    }
                    //MANAGED ASSET SUMMARY
                    else if (currentGroup == "{MERGEFIELD BeginGroup:ManagedAssetSummary}")
                    {
                        if (!managedAssetSummary_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Plans}" && activeGroups.Contains("ManagedAssetSummary}"))
                    {
                        if (!managedAssetSummary_Plans_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:CostsAndChargesSummary}" && activeGroups.Contains("ManagedAssetSummary}"))
                    {
                        if (!managedAssetSummary_CostsAndChargesSummary_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {

                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:CostsAndCharges}" && activeGroups.Contains("ManagedAssetSummary}"))
                    {
                        if (!managedAssetSummary_CostsAndCharges_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:SubPlans}" && activeGroups.Contains("ManagedAssetSummary}"))
                    {
                        if (!managedAssetSummary_SubPlans_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:AllPlans}" && activeGroups.Contains("ManagedAssetSummary}"))
                    {
                        if (!managedAssetSummary_AllPlans_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }

                    //INVESTMENTS
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Investments}")
                    {
                        if (!investments_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:PlansWithFunds}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_PlansWithFunds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_FundsWithoutPortfolio_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Funds}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_Funds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:LookThrough}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_LookThrough_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Assets}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_Assets_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundsWithPortfolio}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_FundsWithPortfolio_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Portfolios}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_Portfolios_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:SubPlansWithFunds}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_SubPlansWithFunds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:CostsAndChargesSummary}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_CostsAndChargesSummary_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:CostsAndCharges}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_CostsAndCharges_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:AssetClasses}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_AssetClasses_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Summarised}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_Summarised_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Detailed}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_Detailed_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Regions}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_Regions_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:SubPlansWithoutFunds}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_SubPlansWithoutFunds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:PlansWithoutFunds}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_PlansWithoutFunds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:AllPlans}" && activeGroups.Contains("Investments}"))
                    {
                        if (!investments_AllPlans_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    //PENSIONS
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Pensions}")
                    {
                        if (!pensions_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:PlansWithFunds}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_PlansWithFunds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:CostsAndChargesSummary}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_CostsAndChargesSummary_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:CostsAndCharges}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_CostsAndCharges_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_FundsWithoutPortfolio_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Funds}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_Funds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:LookThrough}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_LookThrough_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Assets}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_Assets_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundsWithPortfolio}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_FundsWithPortfolio_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Portfolios}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_Portfolios_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:AssetClasses}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_AssetClasses_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Summarised}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_Summarised_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Detailed}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_Detailed_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Regions}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_Regions_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:SubPlansWithFunds}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_SubPlansWithFunds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:SubPlansWithoutFunds}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_SubPlansWithoutFunds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:PlansWithoutFunds}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_PlansWithoutFunds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:AllPlans}" && activeGroups.Contains("Pensions}"))
                    {
                        if (!pensions_AllPlans_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    //General Insurances
                    else if (currentGroup == "{MERGEFIELD BeginGroup:GeneralInsurances}")
                    {
                        if (!generalInsurances_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Plans}" && activeGroups.Contains("GeneralInsurances}"))
                    {
                        if (!generalInsurances_Plans_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}" && activeGroups.Contains("GeneralInsurances}"))
                    {
                        if (!generalInsurances_FundsWithoutPortfolio_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Funds}" && activeGroups.Contains("GeneralInsurances}"))
                    {
                        if (!generalInsurances_Funds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:LookThrough}" && activeGroups.Contains("GeneralInsurances}"))
                    {
                        if (!generalInsurances_LookThrough_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Assets}" && activeGroups.Contains("GeneralInsurances}"))
                    {
                        if (!generalInsurances_Assets_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    //Mortgages
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Mortgages}")
                    {
                        if (!mortgages_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Plans}" && activeGroups.Contains("Mortgages}"))
                    {
                        if (!mortgages_Plans_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}" && activeGroups.Contains("Mortgages}"))
                    {
                        if (!mortgages_FundsWithoutPortfolio_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Funds}" && activeGroups.Contains("Mortgages}"))
                    {
                        if (!mortgages_Funds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:LookThrough}" && activeGroups.Contains("Mortgages}"))
                    {
                        if (!mortgages_LookThrough_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Assets}" && activeGroups.Contains("Mortgages}"))
                    {
                        if (!mortgages_Assets_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    //Protections
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Protections}")
                    {
                        if (!protections_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Plans}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_Plans_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundsWithoutPortfolio}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_FundsWithoutPortfolio_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Funds}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_Funds_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:LookThrough}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_LookThrough_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Assets}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_Assets_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:AssetClasses}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_AssetClasses_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Regions}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_Regions_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Covers}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_Covers_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Summarised}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_Summarised_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Detailed}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_Detailed_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:CoverDetails}" && activeGroups.Contains("Protections}"))
                    {
                        if (!protections_CoverDetails_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    //Tax Wrappers
                    else if (currentGroup == "{MERGEFIELD BeginGroup:TaxWrappers}")
                    {
                        if (!taxWrappers_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Plans}" && activeGroups.Contains("TaxWrappers}"))
                    {
                        if (!taxWrappers_Plans_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    //FundXRay
                    else if (currentGroup == "{MERGEFIELD BeginGroup:FundXRay}")
                    {
                        if (!fundXRay_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:AssetClasses}" && activeGroups.Contains("FundXRay}"))
                    {
                        if (!fundXRay_AssetClasses_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Regions}" && activeGroups.Contains("FundXRay}"))
                    {
                        if (!fundXRay_Regions_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Summarised}" && activeGroups.Contains("FundXRay}"))
                    {
                        if (!fundXRay_Summarised_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                    else if (currentGroup == "{MERGEFIELD BeginGroup:Detailed}" && activeGroups.Contains("FundXRay}"))
                    {
                        if (!fundXRay_Detailed_Group.Contains(field))
                        {
                            File.AppendAllText(logFile, field + " -- Could not be found in this path" + Environment.NewLine);
                            failedMergefields_Hierarchy.Add(field);
                        }
                        else
                        {
                            File.AppendAllText(logFile, field + Environment.NewLine);
                        }
                    }
                }
            }
            return failedMergefields_Hierarchy;
        }
    }
}







