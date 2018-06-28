namespace SPOSiteProvisioningFunctions.Model
{
    public class UpdateMetadataRequest
    {
        public string Type { get; set; }
        public UpdateMetadataJob[] Sites { get; set; }
    }
}
