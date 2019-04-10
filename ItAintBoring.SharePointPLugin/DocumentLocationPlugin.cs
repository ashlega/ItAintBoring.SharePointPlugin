using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using System.Text;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Xml;
using System.Linq;

namespace ItAintBoring.SharePointPlugin
{
    public class DocumentLocationPlugin: IPlugin
    {
        SharePointClient cli = null;
        public DocumentLocationPlugin(string unsecureConfig, string secureConfig)
        {
            XmlDocument doc = new XmlDocument();

            /*
             
             <settings>
                 <clientId></clientId>
                 <clientKey></clientKey>
                 <tenantId></tenantId>
                 <siteRoot></siteRoot>
             </settings>
             
             * */

            doc.LoadXml(secureConfig);
            var settings = doc.SelectSingleNode("settings");
            string clientId = settings.SelectSingleNode("clientId").InnerText;
            string clientKey = settings.SelectSingleNode("clientKey").InnerText;
            string tenantId = settings.SelectSingleNode("tenantId").InnerText;
            string siteRoot = settings.SelectSingleNode("siteRoot").InnerText;

            cli = new SharePointClient(clientId, clientKey, tenantId, siteRoot);
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                ITracingService tracingService =
                    (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                IPluginExecutionContext context = (IPluginExecutionContext)
                    serviceProvider.GetService(typeof(IPluginExecutionContext));

                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                Entity target = (Entity)context.InputParameters["Target"];
                string recordFolder = target.Id.ToString().Replace("{ ", "").Replace("}", "");
                string relativeUrl = "DynamicsDocs/" + recordFolder;
                

                QueryExpression qe = new QueryExpression("sharepointdocumentlocation");
                qe.ColumnSet = new ColumnSet("sharepointdocumentlocationid");
                qe.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, "DynamicsDocs"));
                var parentLocation = service.RetrieveMultiple(qe).Entities.FirstOrDefault();

                if (parentLocation != null)
                {
                    cli.RunQuery("folders", "{ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': '" + relativeUrl + "'}");

                    Entity location = new Entity("sharepointdocumentlocation");
                    location["name"] = target.Id.ToString();
                    location["parentsiteorlocation"] = parentLocation.ToEntityReference();
                    location["relativeurl"] = recordFolder;
                    location["regardingobjectid"] = target.ToEntityReference();
                    service.Create(location);
                }

                

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message + ex.StackTrace);
            }
        }
    }
}
