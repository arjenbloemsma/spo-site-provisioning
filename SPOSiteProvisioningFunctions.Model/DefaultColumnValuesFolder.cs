using System.Collections.Generic;

namespace SPOSiteProvisioningFunctions.Model
{
    public class DefaultColumnValuesFolder
    {
        /// <summary>
        /// The relative path to the folder
        /// </summary>
        public string Path { get; set; }
        /// <summary>
        /// A list of default column value pairs
        /// </summary>
        public List<DefaultColumnValue> DefaultColumnValues { get; set; }
    }
}