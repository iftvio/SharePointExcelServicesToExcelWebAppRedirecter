using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Collections.ObjectModel;

namespace SPExcelWebAppRedirecter.Features.Feature1
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// https://spmatt.wordpress.com/2013/05/22/using-spwebconfigmodification-to-update-the-web-config-in-sharepoint-2013/
    /// https://blogs.msdn.microsoft.com/jjameson/2010/03/31/sharepoint-features-activated-by-default/
    /// https://msdn.microsoft.com/en-us/library/ms436075.aspx
    /// </remarks>

    [Guid("23d91ccd-7e4a-4464-9f20-efc3ab15d274")]
    public class Feature1EventReceiver : SPFeatureReceiver
    {
        // Uncomment the method below to handle the event raised after a feature has been activated.

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            SPWebApplication webApp = properties.Feature.Parent as SPWebApplication;
            RemoveAllCustomisations(webApp);

            #region Enable session state

            SPWebConfigModification httpRuntimeModification = new SPWebConfigModification();
            httpRuntimeModification.Path = "configuration/system.webServer/modules";
            httpRuntimeModification.Name = "add[@name='ExcelWebAppRedirecter']";
            httpRuntimeModification.Sequence = 0;
            httpRuntimeModification.Owner = "SPExcelWebAppRedirecter";
            httpRuntimeModification.Type = SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode;
            httpRuntimeModification.Value = "<add name=\"ExcelWebAppRedirecter\" type=\"ExcelWebAppRedirecter.Redirecter, ExcelWebAppRedirecter, Version=1.0.0.0, Culture=neutral, PublicKeyToken=ab1b281a9b9b52f9\" />";
            webApp.WebConfigModifications.Add(httpRuntimeModification);

            #endregion

            /*Call Update and ApplyWebConfigModifications to save changes*/
            webApp.Update();
            webApp.Farm.Services.GetValue<SPWebService>().ApplyWebConfigModifications();
        }


        // Uncomment the method below to handle the event raised before a feature is deactivated.

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPWebApplication webApp = properties.Feature.Parent as SPWebApplication;
            RemoveAllCustomisations(webApp);
        }


        // Uncomment the method below to handle the event raised after a feature has been installed.

        //public override void FeatureInstalled(SPFeatureReceiverProperties properties)
        //{
        //}


        // Uncomment the method below to handle the event raised before a feature is uninstalled.

        public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        {
            Microsoft.SharePoint.Administration.SPFarm Farm = Microsoft.SharePoint.Administration.SPFarm.Local;
            foreach (SPService Service in Farm.Services)
            {
                if (Service is SPWebService)
                {
                    Microsoft.SharePoint.Administration.SPWebService SPWebService=(Microsoft.SharePoint.Administration.SPWebService)Service;
                    foreach(SPWebApplication webApp in SPWebService.WebApplications)
                    {
                        RemoveAllCustomisations(webApp);
                    }
                }
            }
        }

        // Uncomment the method below to handle the event raised when a feature is upgrading.

        //public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, System.Collections.Generic.IDictionary<string, string> parameters)
        //{
        //}

        private void RemoveAllCustomisations(SPWebApplication webApp)
        {
            if (webApp != null)
            {
                Collection<SPWebConfigModification> collection = webApp.WebConfigModifications;
                int iStartCount = collection.Count;

                // Remove any modifications that were originally created by the owner.
                for (int c = iStartCount - 1; c >= 0; c--)
                {
                    SPWebConfigModification configMod = collection[c];

                    if (configMod.Owner == "SPExcelWebAppRedirecter")
                    {
                        collection.Remove(configMod);
                    }
                }

                // Apply changes only if any items were removed.
                if (iStartCount > collection.Count)
                {
                    webApp.Update();
                    webApp.Farm.Services.GetValue<SPWebService>().ApplyWebConfigModifications();
                }
            }
        }
    }
}
