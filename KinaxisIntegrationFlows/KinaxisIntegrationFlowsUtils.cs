using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using SimioAPI;
using SimioAPI.Extensions;
using System.Text.RegularExpressions;

namespace KinaxisIntegrationFlows
{
    internal class KinaxisIntegrationFlowsUtils
    {
        internal static int runExport(string exporterName, ref ITable table, bool interactiveResults, ref string errorMessage)
        {
            foreach (var exportBinder in table.ExportBindings)
            {
                if (exportBinder.Name.StartsWith(exporterName))
                {
                    table.ExportBindings.ActiveExportBinding = exportBinder;
                    IDataExportResult results = null;
                    if (interactiveResults) results = table.ExportForInteractive();
                    else results = table.ExportForPlan();
                    if (results.Complete == false)
                    {
                        errorMessage = table.Name + " Export Error :" + results.Message;
                        return -1;
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            errorMessage = exporterName + " exporter not found on " + table.Name + " table";
            return -1;
        }

        internal static int runImport(string importerName, ref ITable table, ref string errorMessage)
        {
            foreach (var importBinder in table.ImportBindings)
            {
                if (importBinder.Name.StartsWith(importerName))
                {
                    table.ImportBindings.ActiveImportBinding = importBinder;
                    var results = table.Import();
                    if (results.Complete == false)
                    {
                        errorMessage = table.Name + " Import Error :" + results.Message;
                        return -1;
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            errorMessage = importerName + " importer not found on " + table.Name + " table";
            return -1;
        }

        internal static int runImportFlowForConnectTablAndResultsTable(IModel model, string importerName, string queryIDTableName, string queryIDTablePropertyName, string importTableName)
        {
            string errorMessage = String.Empty;
            var table = model.Tables[queryIDTableName];
            if (table == null) throw new Exception(queryIDTableName + " table not found");
            else
            {
                int returnValue = KinaxisIntegrationFlowsUtils.runImport(importerName, ref table, ref errorMessage);
                if (returnValue == -1) throw new Exception(errorMessage);
                else
                {
                    if (table.Rows.Count == 0) throw new Exception(queryIDTableName + " no results found");
                    else
                    {
                        model.Properties[queryIDTablePropertyName].Value = Regex.Replace(table.Rows[0].Properties[0].Value, "#", "%23");
                        table = model.Tables[importTableName];
                        if (table == null) throw new Exception(importTableName + " table not found");
                        else
                        {
                            errorMessage = String.Empty;
                            returnValue = KinaxisIntegrationFlowsUtils.runImport(importerName, ref table, ref errorMessage);
                            if (returnValue == -1) throw new Exception(errorMessage);
                            else
                            {
                                return returnValue;
                            }
                        }
                    }
                }
            }
        }

        internal static int runImportFlowForOrdersConnectTablAndResultsTable(IModel model, string importerName, string queryIDTableName, string queryIDDemandTablePropertyName, string importDemandTableName, string queryIDSupplyTablePropertyName, string importSupplyTableName, bool runSupply)
        {
            string errorMessage = String.Empty;
            var table = model.Tables[queryIDTableName];
            if (table == null) throw new Exception(queryIDTableName + " table not found");
            else
            {
                int returnValue = KinaxisIntegrationFlowsUtils.runImport(importerName, ref table, ref errorMessage);
                if (returnValue == -1) throw new Exception(errorMessage);
                else
                {
                    if (table.Rows.Count < 2) throw new Exception(queryIDTableName + " not enough results found.  Less than 2 row found");
                    else
                    {
                        if (table.Rows[0].Properties[0].Value != "Demand Orders") throw new Exception(queryIDDemandTablePropertyName + " not found");
                        model.Properties[queryIDDemandTablePropertyName].Value = Regex.Replace(table.Rows[0].Properties[1].Value, "#", "%23");
                        if (table.Rows[1].Properties[0].Value != "Supply Orders") throw new Exception(queryIDSupplyTablePropertyName + " not found");
                        model.Properties[queryIDSupplyTablePropertyName].Value = Regex.Replace(table.Rows[1].Properties[1].Value, "#", "%23");
                        table = model.Tables[importDemandTableName];
                        if (table == null) throw new Exception(importDemandTableName + " table not found");
                        else
                        {
                            errorMessage = String.Empty;
                            returnValue = KinaxisIntegrationFlowsUtils.runImport(importerName, ref table, ref errorMessage);
                            if (returnValue == -1) throw new Exception(errorMessage);
                        }
                        // exit if false
                        if (runSupply == false) return returnValue;
                        table = model.Tables[importSupplyTableName];
                        if (table == null) throw new Exception(importSupplyTableName + " table not found");
                        else
                        {
                            errorMessage = String.Empty;
                            returnValue = KinaxisIntegrationFlowsUtils.runImport(importerName, ref table, ref errorMessage);
                            if (returnValue == -1) throw new Exception(errorMessage);
                            else
                            {
                                return returnValue;
                            }
                        }
                    }
                }
            }
        }

        internal static int runExportFlowForConnectTablAndResultsTable(IModel model, string importerName, string exporterName, bool interactiveResults, string queryIDTableName, string queryIDTablePropertyName, string importTableName)
        {
            string errorMessage = String.Empty;
            var table = model.Tables[queryIDTableName];
            if (table == null) throw new Exception(queryIDTableName + " table not found");
            else
            {
                int returnValue = KinaxisIntegrationFlowsUtils.runImport(importerName, ref table, ref errorMessage);
                if (returnValue == -1) throw new Exception(errorMessage);
                else
                {
                    if (table.Rows.Count == 0) throw new Exception(queryIDTableName + " no results found");
                    else
                    {
                        model.Properties[queryIDTablePropertyName].Value = Regex.Replace(table.Rows[0].Properties[0].Value, "#", "%23");
                        table = model.Tables[importTableName];
                        if (table == null) throw new Exception(importTableName + " table not found");
                        else
                        {
                            errorMessage = String.Empty;
                            returnValue = KinaxisIntegrationFlowsUtils.runExport(exporterName, ref table, interactiveResults, ref errorMessage);
                            if (returnValue == -1) throw new Exception(errorMessage);
                            else
                            {
                                return returnValue;
                            }
                        }
                    }
                }
            }
        }
    }
}