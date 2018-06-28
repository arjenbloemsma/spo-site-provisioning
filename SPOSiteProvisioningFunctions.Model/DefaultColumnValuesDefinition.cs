using System;
using System.Collections.Generic;

namespace SPOSiteProvisioningFunctions.Model
{
    /// <summary>
    /// Object to hold the definition for the default column values.
    /// </summary>
    public class DefaultColumnValuesDefinition
    {
        private List<DefaultColumnValuesLibrary> _libraryCollection;

        /// <summary>
        /// A list containg a the libaries on which default column values must be set.
        /// </summary>
        public List<DefaultColumnValuesLibrary> Libraries {
            get { return _libraryCollection; }
            set { _libraryCollection = value; }
        }

        /// <summary>
        /// A property which indicates when this definition was applied.
        /// </summary>
        public DateTime? AppliedOn { get; set; }

        public DefaultColumnValuesDefinition()
        {
            _libraryCollection = new List<DefaultColumnValuesLibrary>();
        }
    }
}
