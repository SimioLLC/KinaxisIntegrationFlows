using SimioAPI;
using SimioAPI.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace KinaxisIntegrationFlows
{
    internal class KinaxisIntegrationImportFlowInteractive : IDesignAddIn, IDesignAddInGuiDetails
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
                int returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToParts", "PartsQueryID", "Parts");
                if (returnValue == 0) returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToParts", "PartsQueryID", "Sites");
                if (returnValue == 0) returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToParts", "PartsQueryID", "PartsSites");
                if (returnValue == 0) returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToBOM", "BOMQueryID", "BOM");
                if (returnValue == 0) returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForOrdersConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToOrders", "DemandOrdersQueryID", "DemandOrders", "SupplyOrdersQueryID", "SupplyOrders", true);
                if (returnValue == 0) returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForOrdersConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToOrders", "DemandOrdersQueryID", "DemandOrderDetails", "SupplyOrdersQueryID", "SupplyOrders", false);
                if (returnValue == 0) MessageBox.Show("Import Flow Completed", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);         
            }
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

    internal class KinaxisIntegrationExportFlowInteractive : IDesignAddIn, IDesignAddInGuiDetails
    {
        #region IDesignAddIn Members

        /// <summary>
        /// Property returning the name of the add-in. This name may contain any characters and is used as the display name for the add-in in the UI.
        /// </summary>
        public string Name
        {
            get { return "Export Flow"; }
        }

        /// <summary>
        /// Property returning a short description of what the add-in does.
        /// </summary>
        public string Description
        {
            get { return "Export Flow with Interactive Resutls."; }
        }

        /// <summary>
        /// Property returning an icon to display for the add-in in the UI.
        /// </summary>
        public System.Drawing.Image Icon
        {
            get { return Properties.Resources.ExportFlow; }
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
                string exporterName = "CSV Data Exporter1";
                int returnValue = KinaxisIntegrationFlowsUtils.runExportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, exporterName, true, "ConnectToParts", "PartsQueryID", "DemandOrderOutputs");
            }
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

    internal class KinaxisIntegrationImportFlowPlan : IDesignAddIn, IDesignAddInGuiDetails
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
                int returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToParts", "PartsQueryID", "Parts");
                if (returnValue == 0) returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToParts", "PartsQueryID", "Sites");
                if (returnValue == 0) returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToParts", "PartsQueryID", "PartsSites");
                if (returnValue == 0) returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToBOM", "BOMQueryID", "BOM");
                if (returnValue == 0) returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForOrdersConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToOrders", "DemandOrdersQueryID", "DemandOrders", "SupplyOrdersQueryID", "SupplyOrders", true);
                if (returnValue == 0) returnValue = KinaxisIntegrationFlowsUtils.runImportFlowForOrdersConnectTablAndResultsTable(context.ActiveModel, importerName, "ConnectToOrders", "DemandOrdersQueryID", "DemandOrderDetails", "SupplyOrdersQueryID", "SupplyOrders", false);
                if (returnValue == 0) MessageBox.Show("Import Flow Completed", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        #endregion

        #region IDesignAddInGuiDetails Members

        public string CategoryName
        {
            get { return "Planning Table Tools"; }
        }

        public string TabName
        {
            get { return "Content"; }
        }

        public string GroupName
        {
            get { return "Kinaxis"; }
        }

        #endregion
    }

    internal class KinaxisIntegrationFlowPlan : IDesignAddIn, IDesignAddInGuiDetails
    {
        #region IDesignAddIn Members

        /// <summary>
        /// Property returning the name of the add-in. This name may contain any characters and is used as the display name for the add-in in the UI.
        /// </summary>
        public string Name
        {
            get { return "Export Flow"; }
        }

        /// <summary>
        /// Property returning a short description of what the add-in does.
        /// </summary>
        public string Description
        {
            get { return "Export Flow with Plan Resutls."; }
        }

        /// <summary>
        /// Property returning an icon to display for the add-in in the UI.
        /// </summary>
        public System.Drawing.Image Icon
        {
            get { return Properties.Resources.ExportFlow; }
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
                string exporterName = "CSV Data Exporter1";
                int returnValue = KinaxisIntegrationFlowsUtils.runExportFlowForConnectTablAndResultsTable(context.ActiveModel, importerName, exporterName, false, "ConnectToParts", "PartsQueryID", "DemandOrderOutputs");
            }
        }

        #endregion

        #region IDesignAddInGuiDetails Members

        public string CategoryName
        {
            get { return "Planning Table Tools"; }
        }

        public string TabName
        {
            get { return "Content"; }
        }

        public string GroupName
        {
            get { return "Kinaxis"; }
        }

        #endregion
    }

}
