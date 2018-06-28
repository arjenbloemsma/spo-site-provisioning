using System;

namespace SPOSiteProvisioningFunctions.Model
{
    public static class DefaultColumnValuesExtensions
    {
        /// <summary>
        /// Checks if the current default column defnition has been applied to this site by validation if a time stamp is present
        /// </summary>
        /// <param name="definition">The definition object</param>
        /// <returns>True if a time stamp is present, otherwise false</returns>
        public static bool HasBeenApplied(this DefaultColumnValuesDefinition definition)
        {
            return definition.AppliedOn.HasValue;
        }

        /// <summary>
        /// Add a time stamp containing the current time
        /// </summary>
        /// <param name="definition">The definition object</param>
        /// <returns>The updated definition object</returns>
        public static DefaultColumnValuesDefinition AddTimeStampToDefinition(this DefaultColumnValuesDefinition definition)
        {
            definition.AppliedOn = DateTime.Now;
            return definition;
        }
    }
}