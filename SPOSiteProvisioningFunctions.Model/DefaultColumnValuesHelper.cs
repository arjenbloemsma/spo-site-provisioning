using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Schema;
using OfficeDevPnP.Core.Entities;
using System.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace SPOSiteProvisioningFunctions.Model
{
    public static class DefaultColumnValuesHelper
    {
        /// <summary>
        /// Validates if the given JSON string is a valid Default Column Values definition object
        /// </summary>
        /// <param name="definitionJson">The default column values definition object</param>
        /// <param name="errorsMessages">The list of errors, if any</param>
        /// <returns>True if valid, otherwise false</returns>
        public static bool IsValidDefinition(string definitionJson, out IEnumerable<string> errorsMessages)
        {
            bool result = false;

            string fileLocation = ConfigurationManager.AppSettings["mar:defaultColumnValuesSchema"];
            IList<ValidationError> errors;

            // Validate if the JSON is valid by checking it against the schema
            using (JsonTextReader reader = new JsonTextReader(new StringReader(
                System.IO.File.OpenText(fileLocation).ReadToEnd())))
                {
                    JSchema schema = JSchema.Load(reader);
                    JToken json = JObject.Parse(definitionJson);
                    result = json.IsValid(schema, out errors);
                }

            errorsMessages = errors.Select(e => $"{e.LineNumber}: {e.Message}");

            return result;
        }

        /// <summary>
        /// Creates a DefaultColumnTextValue based on the given parameters
        /// </summary>
        /// <param name="folderPath">The relative folder path</param>
        /// <param name="field">The field</param>
        /// <param name="value">The value</param>
        /// <returns></returns>
        public static IDefaultColumnValue ToDefaultColumnTextValueValue(string folderPath, Field field, string value)
        {
            if (folderPath == null)
            {
                throw new ArgumentNullException(nameof(folderPath));
            }
            if (field == null)
            {
                throw new ArgumentNullException(nameof(field));
            }

            return new DefaultColumnTextValue()
            {
                FolderRelativePath = folderPath,
                FieldInternalName = field.InternalName,
                Text = value
            };
        }

        /// <summary>
        /// Returns the library based on the given name
        /// </summary>
        /// <param name="web">The web</param>
        /// <param name="libraryName">The name of the library</param>
        /// <returns>The library or an exception if the list couldn't be found or isn't of the library type</returns>
        public static List GetLibrary(Web web, string libraryName)
        {
            if (string.IsNullOrWhiteSpace(libraryName) == true)
            {
                throw new ArgumentNullException(nameof(libraryName));
            }

            var list = web.GetListByTitle(libraryName);
            if (list == null)
            {
                throw new ArgumentException($"List {libraryName} coundn't be found.");
            }

            if (list.BaseTemplate != (int) ListTemplateType.DocumentLibrary &&
                list.BaseTemplate != (int) ListTemplateType.WebPageLibrary &&
                list.BaseTemplate != (int) ListTemplateType.PictureLibrary &&
                list.BaseTemplate != 851) // Asset library
            {
                throw new ArgumentException($"List {libraryName} is not a document library.");
            }

            return list;
        }

        /// <summary>
        /// Returns the field based on the given name
        /// </summary>
        /// <param name="ctx">The client context</param>
        /// <param name="list">The list</param>
        /// <param name="fieldName">The inyternal name or title of the field</param>
        /// <returns>The field</returns>
        public static Field GetField(ClientContext ctx, List list, string fieldName)
        {
            if (list == null)
            {
                throw new ArgumentNullException(nameof(list));
            }
            if (string.IsNullOrWhiteSpace(fieldName) == true)
            {
                throw new ArgumentNullException(nameof(fieldName));
            }

            var field = list.Fields.GetByInternalNameOrTitle(fieldName);
            if (field == null)
            {
                throw new ArgumentException($"Field {fieldName} coundn't be found.");
            }
            ctx.Load(field);
            ctx.ExecuteQueryRetry();

            return field;
        }

        /// <summary>
        /// Returns the folder based on the given name
        /// </summary>
        /// <param name="ctx">The client context</param>
        /// <param name="list">The list</param>
        /// <param name="name">The name of the folder</param>
        /// <returns>The folder</returns>
        public static Folder GetFolder(ClientContext ctx, List list, string name)
        {
            if (list == null)
            {
                throw new ArgumentNullException(nameof(list));
            }
            if (string.IsNullOrWhiteSpace(name) == true)
            {
                throw new ArgumentNullException(nameof(name));
            }

            Folder folder = null;
            folder = name == "/" ? list.RootFolder : list.RootFolder.ResolveSubFolder(name);
            if (folder == null)
            {
                throw new ArgumentException($"Folder {name} couldn't be found.");
            }
            ctx.Load(folder, f => f.Name, f => f.ServerRelativeUrl);
            ctx.ExecuteQueryRetry();

            return folder;
        }
    }
}