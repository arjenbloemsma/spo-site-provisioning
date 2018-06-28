using Marlink.SharePoint.Provisioning.Service.Common;
using Marlink.SharePoint.Provisioning.Service.Common.Security;
using Marlink.SharePoint.Provisioning.Service.Common.SharePoint;
using Microsoft.Azure;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.ServiceBus.Messaging;
using Microsoft.SharePoint.Client;
using SPOSiteProvisioningFunctions.Model;
using System;
using System.Runtime.Serialization.Json;
using System.Threading;

namespace SPOSiteProvisioningFunctions
{
    public static class MonitorSiteCollectionCreation
    {
        private const string FunctionName = "monitor-site-collection-creation";

        [FunctionName(FunctionName)]
        public static async void Run(
            [OrchestrationTrigger] DurableOrchestrationContext orchestrationContext,
            [ServiceBus("new-sites-topic", Connection = "ManageTopicConnection")]ICollector<BrokeredMessage> newSitesTopic,
            TraceWriter log)
        {
            log.Info($"C# Orchestration trigger function '{FunctionName}' started.");

            bool siteHasBeenCreated = false;
            var siteCollectionCreationData = orchestrationContext.GetInput<MonitorSiteCollectionCreationData>();
            log.Info($"Monitor creation of site '{siteCollectionCreationData.FullSiteUrl}'.");
            if (siteCollectionCreationData.TimeStamp == null
                || DateTime.Now.Subtract(siteCollectionCreationData.TimeStamp) > new TimeSpan(24, 10, 00))
            {
                log.Info($"SiteCollection {siteCollectionCreationData.FullSiteUrl} was not created within 24 hours or timestamp was empty.");
                return;
            }

            var clientContextManager = new ClientContextManager(new BaseConfiguration(), new CertificateManager());
            using (var adminCtx = clientContextManager.GetAzureADAppOnlyAuthenticatedContext(CloudConfigurationManager.GetSetting("TenantAdminUrl")))
            {
                var tenant = new Tenant(adminCtx);
                try
                {
                    //Get the site name
                    var properties = tenant.GetSitePropertiesByUrl(siteCollectionCreationData.FullSiteUrl, false);
                    tenant.Context.Load(properties);
                    // Will cause an exception if site URL is not there. Not optimal, but the way it works.
                    tenant.Context.ExecuteQueryRetry();
                    log.Info($"Site creation status: '{properties.Status}'.");
                    siteHasBeenCreated = properties.Status.Equals("Active", StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    try
                    {
                        // Check if a site collection with this URL has been recycled (exists in garbage bin)
                        var deletedProperties = tenant.GetDeletedSitePropertiesByUrl(siteCollectionCreationData.FullSiteUrl);
                        tenant.Context.Load(deletedProperties);
                        tenant.Context.ExecuteQueryRetry();
                        if (deletedProperties.Status.Equals("Recycled", StringComparison.OrdinalIgnoreCase))
                        {
                            log.Info($"SiteCollection with URL{siteCollectionCreationData.FullSiteUrl} already exists in recycle bin.");

                            var provisioningSiteUrl = CloudConfigurationManager.GetSetting("ProvisioningSite");
                            using (var ctx = clientContextManager.GetAzureADAppOnlyAuthenticatedContext("https://mobsatdev.sharepoint.com/sites/site-provisioning"))
                            {
                                // Todo: get list title from configuration.
                                // Assume that the web has a list named "PnPProvisioningJobs". 
                                List provisioningJobsList = ctx.Web.Lists.GetByTitle("PnPProvisioningJobs");
                                ListItem listItem = provisioningJobsList.GetItemById(siteCollectionCreationData.ListItemID);
                                // Write a new value to the PnPProvisioningJobStatus field of
                                // the PnPProvisioningJobs item.
                                listItem["PnPProvisioningJobStatus"] = "Failed (site with same URL exists in recycle bin)"; //ToDo: Should we have different statusses for actual site creation and the applying of the template?
                                listItem.Update();

                                ctx.ExecuteQuery();
                            }
                            return;
                        }
                    } catch
                    {
                        siteHasBeenCreated = false;
                    }
                }

                    //Check if provisioning of the SiteCollection is complete.
                while (!siteHasBeenCreated)
                {
                    //Wait for 1 minute and then try again
                    DateTime deadline = orchestrationContext.CurrentUtcDateTime.Add(TimeSpan.FromMinutes(1));
                    await orchestrationContext.CreateTimer(deadline, CancellationToken.None);
                }

                if (siteHasBeenCreated)
                {
                    log.Info($"SiteCollection {siteCollectionCreationData.FullSiteUrl} created.");

                    var createSiteCollectionMsg = new BrokeredMessage(siteCollectionCreationData.CreateSiteCollectionJob,
                        new DataContractJsonSerializer(typeof(CreateSiteCollectionJob)))
                    {
                        ContentType = "application/json",
                        Label = "ApplyTemplate"
                    };
                    newSitesTopic.Add(createSiteCollectionMsg);
                }
            }
        }
    }
}