using Marlink.SharePoint.Provisioning.Service.Common;
using Marlink.SharePoint.Provisioning.Service.Common.Security;
using Marlink.SharePoint.Provisioning.Service.Common.SharePoint;
using Microsoft.Azure;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.ServiceBus.Messaging;
using Microsoft.SharePoint.Client;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json;
using SPOSiteProvisioningFunctions.Model;
using System;

namespace SPOSiteProvisioningFunctions
{
    public static class UpdateSiteMetadata
    {
        private const string FunctionName = "update-site-metadata";

        [FunctionName(FunctionName)]
        public static void Run(
            [ServiceBusTrigger("site-updates-topic", "update-metadata-subscription", AccessRights.Manage, Connection = "ManageTopicConnection")]BrokeredMessage updateMsg,
            TraceWriter log)
        {
            log.Info($"C# ServiceBus trigger function '{FunctionName}' processed message: {updateMsg.MessageId} (Label': {updateMsg.Label}')");

            var somethingWentWrong = false;
            var clientContextManager = new ClientContextManager(new BaseConfiguration(), new CertificateManager());
            // Todo: validate that URL exists
            var updateMetadataJob = updateMsg.GetBody<UpdateMetadataJob>();
            using (var ctx = clientContextManager.GetAzureADAppOnlyAuthenticatedContext(updateMetadataJob.Url))
            {
                // ToDo; currently we only support updating the Title (incl. Maritime Installations)
                // Maybe use switch state ment here. Come up with some better way of resolving
                // what kind of site we're dealing with (info from property bag or Azure storage table)
                if (updateMetadataJob.Url.ToLowerInvariant().Contains("/inst-") == true)
                {
                    try
                    {
                        const string propertyBagDefaultColumnValues = "_marlink_defaultcolumnvalues";
                        ctx.Web.Title = updateMetadataJob.Title;
                        var definitionJson = ctx.Web.GetPropertyBagValueString(propertyBagDefaultColumnValues, String.Empty);
                        if (string.IsNullOrEmpty(definitionJson))
                        {
                            log.Info($"Definition in site {updateMetadataJob.Url} was empty.");
                            return;
                        }

                        var definition = JsonConvert.DeserializeObject<DefaultColumnValuesDefinition>(definitionJson);
                        const string siteColumnName = "dc_Site";
                        definition.Libraries.ForEach(l =>
                        {
                            l.Folders.ForEach(f =>
                            {
                                var siteColumnIndex = f.DefaultColumnValues.FindIndex(c => String.CompareOrdinal(c.Name, siteColumnName) == 0);
                                if (siteColumnIndex != -1)
                                {
                                    f.DefaultColumnValues[siteColumnIndex].Value = updateMetadataJob.Title;
                                }
                                else
                                {
                                    f.DefaultColumnValues.Add(new DefaultColumnValue()
                                    {
                                        Name = siteColumnName,
                                        Value = updateMetadataJob.Title
                                    });
                                }
                            });
                        });

                        // Save the template information in the target site
                        var updatedDefinition = definition;
                        updatedDefinition.AppliedOn = null;
                        string updatedJson = JsonConvert.SerializeObject(updatedDefinition);
                        // Validate if the defintion confirms to the schema
                        /*
                        if (DefaultColumnValuesHelper.IsValidDefinition(updatedJson, out IEnumerable<string> errors) == false)
                        {
                            log.Error($"Invalid JSON Definition was provided for {updateMetadataJob.Url}.");
                            foreach (var errMsg in errors)
                            {
                                log.Info(errMsg);
                            }
                            return;
                        }
                        */

                        ctx.Web.SetPropertyBagValue("_marlink_defaultcolumnvalues", updatedJson);
                        log.Info($"Updated Default Column Values Definition in site {updateMetadataJob.Url}.");
                    }
                    catch (Exception ex)
                    {
                        log.Error($"Error occured while setting default column values to {updateMetadataJob.Url}.", ex);
                        somethingWentWrong = true;
                    }
                }

                try
                {

                    CloudStorageAccount storageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("AzureWebJobsStorage"));
                    CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
                    CloudTable customerDocumentCenterSitesTable = tableClient.GetTableReference(CloudConfigurationManager.GetSetting("SitesTable"));

                    // ToDo: update Title (and other fields) in Azure Storage
                    var item = new CustomerDocumentCenterSitesTableEntry
                    {
                        PartitionKey = updateMetadataJob.Type,
                        RowKey = updateMetadataJob.ID,
                        ETag = "*",
                        Title = updateMetadataJob.Title,
                    };
                    if (updateMetadataJob.Type == "INST")
                    {
                        item.Site = updateMetadataJob.Title;
                    }
                    var operation = TableOperation.Merge(item);
                    customerDocumentCenterSitesTable.ExecuteAsync(operation);
                }
                catch (Exception ex)
                {
                    log.Error($"Error occured while updating table entry of {updateMetadataJob.Url}.", ex);
                    somethingWentWrong = true;
                }

                if (somethingWentWrong == false) {
                    ctx.Web.Title = updateMetadataJob.Title;
                    ctx.Web.Update();
                    ctx.ExecuteQuery();
                    log.Info($"Updated title of site collection '{updateMetadataJob.Url}' to '{updateMetadataJob.Title}'");
                }
            }
        }
    }
}