namespace SPOSiteProvisioningFunctions.Model
{
    public class UpdateSiteRequest
    {
        public string Type { get; set; }
        public UpdateSiteJob[] Sites { get; set; }
    }
}
