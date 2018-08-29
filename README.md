# spo-site-provisioning
Function app for provisioning SharePoint Online site collections and applying templates and updates to them.

## Table of Contents
- [Description](#description)
- [Documentation](#documentation)
- [Dependencies](#dependencies)
- [Technology](#technology)
- [CI/CD](#cicd)
- [Deployment](#deployment)
- [Testing](#testing)
- [Functionality](#functionality)
- [Project Team](#project-team)

## Description
These functions make the SharePoint Online site collection provisioning process event driven.

SharePoint Online site collections are provisioned using a fork of the [PnP provisioning engine](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/introducing-the-pnp-provisioning-engine "Introducing the PnP provisioning engine") which can be found [here](https://spo-siteprovisioning.azurewebsites.net). The provisioning engine will create a .job file and will place that file in the [PnPProvisioningJobs list](https://mobsat.sharepoint.com/sites/site-provisioning/PnPProvisioningJobs). This will trigger a Microsoft Flow which in turn will call the first Azure Function in the process.

For updating site collections, this could be a simple update that involves an adjustment of some metadata or a more complext update like applying an updated template, a POST request must be made to an endpoint. This endpoint expects a JSON message which must contain the info about the update that should take place. One single request can contain update info about one or multiple site collections.

## Documentation
[Modern Site Provisioning and Update Functionality](https://mobsat.sharepoint.com/:u:/s/Documentcenter/EcoZJv5IdM1HuXVT5uZMJZQBX45AU4d8H90UGIRTULzlHg)

## Dependencies
Internal dependencies:

N/A

External dependencies:
+ Azure Servicebus => triggering functions
+ Azure BLOB Storage => storing/reading jobfiles by functions
+ Azure Table Storage => reference data used by functions
+ SharePoint Online =>  creating/updating site collection by functions

## Technology
+ `Visual Studio 2017`
+ `.NET 4.6.1`
+ `Azure Functions`
+ `Azure Servicebus`
+ `Azure Storage`
+ `SharePoint Online`

## CI/CD
VSTS (to be set up)

## Deployment
To run this project locally, just get the sources run, no need to do any additional setup.

## Testing
Description of how to run the tests of this project.

## Functionality
Each of the Azure functions of which the site collection creation and update process consists will be described here.

### AddSiteUpdateRequest
This HTTP triggered function will try to deserialize the message and add the resulting update request to the service bus topic "site-operations-topic".

### ApplyProvisioningTemplate
This function is triggered by messages on the subscription 'apply-template-subscription' of the service bus topic 'site-updates-topic'. The message will contain the name of a job file. The function will retrieve the job file from blob storage and get all required info from it in order to apply the template to the specified site collection. Once done it will update the status in the PnPProvisioningJobs list to 'Running (applying template)'.

### CreateSiteCollection
This function is triggered by messages on the subscription 'create-site-subscription' of the service bus topic 'new-sites-topic'. The message will contain the name of a job file. The function will retrieve the job file from blob storage and get all required info from it in order to start the creation the site collection with the specified metadata. The function will also trigger a durable function [MonitorSiteCollectionCreation](#monitorsitecollectioncreation).

### MonitorSiteCollectionCreation
This is a durable function which is triggered by the function CreateSiteCollection. It will check every minute if the site creation process initiated by the function [CreateSiteCollection](#createsitecollection) has been completed and add a message to the service bus topic 'new-sites-topic' when this is the case. If the site collection is not created within 24 hours it will log a warning message and stop monitoring.

### ProcessNewSiteRequest
This function is triggered by messages on the subscription 'new-sites-subscription' of the service bus topic 'site-operations-topic'. The message will contain the name of a job file. The function will retrieve the job file from the PnPProvisioningJobs list in SharePoint Online and store it in blob storage. It will also place a message in the service bus topic 'new-sites-topic' and set the status of the item PnPProvisioningJobs list to 'Running (creating site collection)'.

### ProcessUpdateSiteRequest
This function is triggered by messages on the subscription 'update-requests-subscription' of the service bus topic 'site-operations-topic'. The function will use the relative URL from this message to retrieve more info about this site collection from the storage table 'CustomerDocumentCenterSites' and create a new message with all this data and place that in the service bus topic 'site-updates-topic'.

### RegisterNewSite
This function is triggered by messages on the subscription 'register-site-subscription' of the service bus topic 'new-sites-topic'. The message will contain the name of a job file. The function will retrieve the job file from blob storage and get all required info from it in order to create an entry in the storage table 'CustomerDocumentCenterSites'.

### SetDefaultColumnValues
This function is triggered by messages on the subscription 'set-default-column-values-subscription' of the service bus topic 'new-sites-topic'. The message will contain the name of a job file. The function will retrieve the job file from blob storage and set the specified template parameters as default column values on the specified site collection. 

### TestSiteCollectionExists
This HTTP triggered function will test if a site collection exists at the (relative) url specified in the query string and return a JSON message with information about the site. The property 'Exists' of this message will always have value and is either true if the site exists or false if it doesn't. Other properties like 'Title' or ‘Relative URL’ will only have a value if the site does exist. Useful in other services that do not have direct access to Sharepoint Online.

### UpdateSiteMetadata
This function is triggered by messages on the subscription 'update-metadata-subscription' of the service bus topic 'site-updates-topic'. The message will contain the metadata to update (currently only Title is supported). The function will retrieve those new values and will update them in all required places; so not only the metadata of the site collection (title, property bag values, default column values), but also in the storage table 'CustomerDocumentCenterSites'.  

### UpdateSiteTemplate
This function is triggered by messages on the subscription 'update-template-subscription' of the service bus topic 'site-updates-topic'. The message will contain references to info required to apply or update the site template. The function will retrieve the template file from SharePoint Online, but only if we don't have it already in Azure Storage as a blob or if that version in Azure Storage is older than one hour. This to make sure that the site collection has the latest version of the provisioning template applied to it, while not constantly downloading the same template file from SharePoint Online.

### ValidateTemplateCache
This HTTP triggered function will try to deserialize the message and update the template specified in the message in the blob storage. This is done by retrieving the template file from SharePoint Online, but only if we don't have it already in Azure Storage as a blob or if that version in Azure Storage is older than one hour. Useful when a new template is uploaded to SharePoint Online by the administrator to make sure that the version of this template in Azure Storage is immediately updated.

## Project Team
Development contributers can be found here: 

`[https://github.com/marlinknl/spo-site-provisioning/graphs/contributors]`
