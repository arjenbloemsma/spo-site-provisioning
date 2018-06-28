using Marlink.SharePoint.Provisioning.Service.Common;
using Marlink.SharePoint.Provisioning.Service.Common.Security;
using Marlink.SharePoint.Provisioning.Service.Common.SharePoint;
using Microsoft.Azure;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.ServiceBus.Messaging;
using Microsoft.SharePoint.Client;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SPOSiteProvisioningFunctions.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SPOSiteProvisioningFunctions
{
    public static class SetDefaultColumnValues
    {
        private const string FunctionName = "set-deafult-column-values";

        [FunctionName(FunctionName)]
        public static void Run(
            [ServiceBusTrigger("new-sites-topic", "set-default-column-values-subscription", AccessRights.Manage, Connection = "ManageTopicConnection")]BrokeredMessage setDefaultColumnValuesMsg,
            TraceWriter log)
        {
            log.Info($"C# Service Bus trigger function '{FunctionName}' processed message: {setDefaultColumnValuesMsg.MessageId} (Label': {setDefaultColumnValuesMsg.Label}')");

            var stream = setDefaultColumnValuesMsg.GetBody<Stream>();
            StreamReader streamReader = new StreamReader(stream);
            string createSiteCollectionJobAsJson = streamReader.ReadToEnd();
            var createSiteCollectionJob = JsonConvert.DeserializeObject<CreateSiteCollectionJob>(createSiteCollectionJobAsJson);

            // ToDo: instead of retrieving this info from blob, get it from table storage
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("AzureWebJobsStorage"));
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            CloudBlobContainer container = blobClient.GetContainerReference(CloudConfigurationManager.GetSetting("JobFilesContainer"));

            var blob = container.GetBlobReference(createSiteCollectionJob.FileNameWithExtension);
            var blobStream = new MemoryStream();
            blob.DownloadToStream(blobStream);
            streamReader = new StreamReader(blobStream);
            blobStream.Position = 0;
            string blobContent = streamReader.ReadToEnd();

            JObject provisioningJobFile = JObject.Parse(blobContent);
            var provisioningTemplateUrl = provisioningJobFile["ProvisioningTemplateUrl"].Value<string>();
            var relativeUrl = provisioningJobFile["RelativeUrl"].Value<string>();

            var tenantUrl = new Uri(CloudConfigurationManager.GetSetting("TenantUrl"));
            Uri.TryCreate(tenantUrl, relativeUrl, out Uri fullSiteUrl);

            // Currently only do this for Maritime Installation sites
            if (relativeUrl.ToLowerInvariant().Contains("/inst-") == false)
            {
                log.Info($"Site collection {fullSiteUrl.AbsoluteUri} is not a maritime Installation site. Skip setting default column values.");
                return;
            }

            // get JSON result objects into a list
            IList<JToken> parameters = provisioningJobFile["TemplateParameters"].Children().ToList();
            // serialize JSON results into .NET objects
            var defaultColumnValues = new List<DefaultColumnValue>();
            foreach (JProperty parameter in parameters)
            {
                defaultColumnValues.Add(new DefaultColumnValue
                {
                    Name = $"dc_{parameter.Name}",
                    Value = parameter.Value.ToObject<string>()
                });
            }

            var folders = new List<DefaultColumnValuesFolder> {
                new DefaultColumnValuesFolder { Path = "/", DefaultColumnValues = defaultColumnValues }
            };

            var definition = new DefaultColumnValuesDefinition();
            definition.Libraries.AddRange(new List<DefaultColumnValuesLibrary> {
                new DefaultColumnValuesLibrary { Name = "Installation", Folders = folders }
                ,new DefaultColumnValuesLibrary { Name = "Logistics", Folders = folders }
                ,new DefaultColumnValuesLibrary { Name = "Network", Folders = folders }
                ,new DefaultColumnValuesLibrary { Name = "Pictures", Folders = folders }
                ,new DefaultColumnValuesLibrary { Name = "Project", Folders = folders }
                ,new DefaultColumnValuesLibrary { Name = "Solutions", Folders = folders }
                ,new DefaultColumnValuesLibrary { Name = "Videos", Folders = folders }
            });

            var clientContextManager = new ClientContextManager(new BaseConfiguration(), new CertificateManager());
            using (var ctx = clientContextManager.GetAzureADAppOnlyAuthenticatedContext(fullSiteUrl.AbsoluteUri))
            {
                try
                {
                    var jsonDefinition = JsonConvert.SerializeObject(definition, Formatting.None);
                    ctx.Web.SetPropertyBagValue("_marlink_defaultcolumnvalues", jsonDefinition);
                    log.Info($"Added Default Column Values Definition object to property bag of site {fullSiteUrl.AbsoluteUri}.");
                }
                catch (Exception ex)
                {
                    log.Error($"Error occured while setting default column values to {fullSiteUrl.AbsoluteUri}.", ex);
                }
            }
        }
    }
}
