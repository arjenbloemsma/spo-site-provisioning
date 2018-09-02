using Microsoft.WindowsAzure.Storage.Table;
using System;

namespace SPOSiteProvisioningFunctions.Model
{
    public class CustomerDocumentCenterSitesTableEntry : TableEntity
    {
        public string Title { get; set; }
        public string URL { get; set; }
        public string CompanyContext { get; set; }
        public string SiteContext { get; set; }
        public string PersonId { get; set; }
        public string InstallationContext { get; set; }
        public string ProjectContext { get; set; }
        public string CallSign { get; set; }
        public string IMO { get; set; }
        public string InstallationId { get; set; }
        public string InstallationType { get; set; }
        public string MMSI { get; set; }
        public string OperationalCustomer { get; set; }
        public string OperationalCustomerID { get; set; }
        public string Site { get; set; }
        public string SiteID { get; set; }
        public int ProvisioningStatus { get; set; }
        public DateTime? Updated { get; set; }
        public string ProvisioningTemplateUrl { get; set; }
    }
}
