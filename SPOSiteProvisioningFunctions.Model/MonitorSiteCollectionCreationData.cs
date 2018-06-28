using System;

namespace SPOSiteProvisioningFunctions.Model
{
    public class MonitorSiteCollectionCreationData
    {
        public int ListItemID { get; set; }
        public string FullSiteUrl { get; set; }
        public DateTime TimeStamp { get; set; }
        public CreateSiteCollectionJob CreateSiteCollectionJob { get; set; }
    }
}
