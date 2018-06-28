using Microsoft.Azure;
using SPOSiteProvisioningFunctions.Common;
using System;

namespace Marlink.SharePoint.Provisioning.Service.Common
{
    public class BaseConfiguration : IConfiguration
    {
        public string GetSetting(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return CloudConfigurationManager.GetSetting(name);
        }
    }
}