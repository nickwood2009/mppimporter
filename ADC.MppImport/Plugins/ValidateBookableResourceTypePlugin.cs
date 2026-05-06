using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ADC.MppImport.Plugins
{
    /// <summary>
    /// Prevents non-User bookable resources from being assigned to project teams.
    /// Register on msdyn_projectteam: Pre-Operation, Create and Update.
    /// </summary>
    public class ValidateBookableResourceTypePlugin : IPlugin
    {
        private const int RESOURCE_TYPE_USER = 3;
        private const string RESOURCE_LOOKUP = "msdyn_bookableresourceid";

        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity target))
                return;

            if (!target.Contains(RESOURCE_LOOKUP))
                return;

            var resourceRef = target.GetAttributeValue<EntityReference>(RESOURCE_LOOKUP);
            if (resourceRef == null)
                return;

            tracingService.Trace("ValidateBookableResourceType: Checking resource {0}", resourceRef.Id);

            var resource = service.Retrieve("bookableresource", resourceRef.Id,
                new ColumnSet("resourcetype", "name"));

            var resourceType = resource.GetAttributeValue<OptionSetValue>("resourcetype");
            var name = resource.GetAttributeValue<string>("name") ?? resourceRef.Id.ToString();

            if (resourceType == null || resourceType.Value != RESOURCE_TYPE_USER)
            {
                tracingService.Trace("ValidateBookableResourceType: Rejected resource '{0}' (type={1})",
                    name, resourceType?.Value);
                throw new InvalidPluginExecutionException(
                    string.Format("Only User-type bookable resources can be assigned to project tasks. " +
                                  "'{0}' is not a User resource.", name));
            }

            tracingService.Trace("ValidateBookableResourceType: Accepted resource '{0}'", name);
        }
    }
}
