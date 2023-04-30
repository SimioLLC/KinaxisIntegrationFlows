using SimioAPI;
using SimioAPI.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace KinaxisIntegrationFlows
{
    internal class KinaxisIntegrationImportFlow : IDesignAddIn, IDesignAddInGuiDetails
    {
        #region IDesignAddIn Members

        /// <summary>
        /// Property returning the name of the add-in. This name may contain any characters and is used as the display name for the add-in in the UI.
        /// </summary>
        public string Name
        {
            get { return "Import Flow"; }
        }

        /// <summary>
        /// Property returning a short description of what the add-in does.
        /// </summary>
        public string Description
        {
            get { return "Import Flow."; }
        }

        /// <summary>
        /// Property returning an icon to display for the add-in in the UI.
        /// </summary>
        public System.Drawing.Image Icon
        {
            get { return Properties.Resources.ImportFlow; }
        }

        /// <summary>
        /// Method called when the add-in is run.
        /// </summary>
        public void Execute(IDesignContext context)
        {
            // This example code places some new objects from the Standard Library into the active model of the project.
            if (context.ActiveModel != null)
            {
                string importerName = "Web API Importer1";
                int returnValue = runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToParts", "PartsQueryID", "Parts");
                if (returnValue == 0) returnValue = runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToParts", "PartsQueryID", "Sites");
                if (returnValue == 0) returnValue = runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToParts", "PartsQueryID", "PartsSites");
                if (returnValue == 0) returnValue = runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToBOM", "BOMQueryID", "BOM");
                if (returnValue == 0) returnValue = runImportFlowForOrdersConnectTablAndResultsTable(context.ActiveModel, importerName,  "ConnectToOrders", "DemandOrdersQueryID", "DemandOrders", "SupplyOrdersQueryID", "SupplyOrders");
                if (returnValue == 0) MessageBox.Show("Import Flow Completed", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);         
            }
        }

        public static int runImportFlowForConnectTablAndResultsTable(IModel model, string importerName, string queryIDTableName, string queryIDTablePropertyName, string importTableName)
        {
            string errorMessage = String.Empty;
            var table = model.Tables[queryIDTableName];
            if (table == null) throw new Exception(queryIDTableName + " table not found");
            else
            {
                int returnValue = runImport(importerName, ref table, ref errorMessage);
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
                            returnValue = runImport(importerName, ref table, ref errorMessage);
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

        public static int runImportFlowForOrdersConnectTablAndResultsTable(IModel model, string importerName, string queryIDTableName, string queryIDDemandTablePropertyName, string importDemandTableName, string queryIDSupplyTablePropertyName, string importSupplyTableName )
        {
            string errorMessage = String.Empty;
            var table = model.Tables[queryIDTableName];
            if (table == null) throw new Exception(queryIDTableName + " table not found");
            else
            {
                int returnValue = runImport(importerName, ref table, ref errorMessage);
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
                            returnValue = runImport(importerName, ref table, ref errorMessage);
                            if (returnValue == -1) throw new Exception(errorMessage);
                        }                      
                        table = model.Tables[importSupplyTableName];
                        if (table == null) throw new Exception(importSupplyTableName + " table not found");
                        else
                        {
                            errorMessage = String.Empty;
                            returnValue = runImport(importerName, ref table, ref errorMessage);
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


        public static int runImport(string importerName, ref ITable table, ref string errorMessage)
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

        #endregion

        #region IDesignAddInGuiDetails Members

        public string CategoryName
        {
            get { return "Table Tools"; }
        }

        public string TabName
        {
            get { return "Kinaxis"; }
        }

        public string GroupName
        {
            get { return "Integration Flows"; }
        }

        #endregion
    }
}
