using System.Collections.Generic;

namespace SPOSiteProvisioningFunctions.Model
{
    public class DefaultColumnValuesLibrary
    {
        /// <summary>
        /// Name of the library on which the default column values should be set
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// A list of folder in which default column values must be set
        /// </summary>
        public List<DefaultColumnValuesFolder> Folders { get; set; }
    }
}