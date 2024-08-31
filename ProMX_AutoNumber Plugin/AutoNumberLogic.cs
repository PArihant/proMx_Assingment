using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Specialized;
using System.Web;

namespace ProMX_AutoNumber
{
    public class AutoNumberLogic : IPlugin
    {
        // Declare the organization service and tracing service variables
        private IOrganizationService _service;
        private ITracingService _traceService;

        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                // Initialize the tracing service for debugging purposes
                _traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

                // Obtain the execution context from the service provider
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

                // Obtain the organization service reference
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                _service = serviceFactory.CreateOrganizationService(context.UserId);

                _traceService.Trace("Start ProMX_AutoNumber Action");

                // Read input parameters
                string recordUrl = (string)context.InputParameters["record"];
                string entity = (string)context.InputParameters["entity"];
                string field = (string)context.InputParameters["field"];
                _traceService.Trace("record_url: " + recordUrl);
                _traceService.Trace("entity: " + entity);
                _traceService.Trace("field: " + field);

                // Extract the record ID from the URL
                Uri uri = new Uri(recordUrl);
                string queryString = uri.Query;
                NameValueCollection queryParameters = HttpUtility.ParseQueryString(queryString);
                string idValue = queryParameters["id"];
                _traceService.Trace("idValue: " + idValue);

                // Retrieve the auto number configuration record for the specified entity
                QueryExpression autoNumberConfigQuery = new QueryExpression("trial_autonumberconfigs")
                {
                    ColumnSet = new ColumnSet("trial_name", "trial_number", "trial_postfix", "trial_prefix", "trial_seedvalue"),
                    Criteria =
                    {
                        Conditions =
                        {
                            new ConditionExpression("statuscode", ConditionOperator.Equal, 1), // Active
                            new ConditionExpression("trial_tablelogicalname", ConditionOperator.Equal, entity)
                        }
                    }
                };

                EntityCollection autoNumberConfigCollection = _service.RetrieveMultiple(autoNumberConfigQuery);
                _traceService.Trace("auto_number_config_collection Count: " + autoNumberConfigCollection.Entities.Count);

                if (autoNumberConfigCollection.Entities.Count > 0)
                {
                    foreach (Entity autoNumberConfigRecord in autoNumberConfigCollection.Entities)
                    {
                        // Retrieve and format the auto number components
                        string displayName = autoNumberConfigRecord.Contains("trial_name") ? (string)autoNumberConfigRecord["trial_name"] : string.Empty;
                        string prefix = autoNumberConfigRecord.Contains("trial_prefix") ? (string)autoNumberConfigRecord["trial_prefix"] : string.Empty;
                        string postfix = autoNumberConfigRecord.Contains("trial_postfix") ? (string)autoNumberConfigRecord["trial_postfix"] : string.Empty;
                        int number = autoNumberConfigRecord.Contains("trial_number") ? (int)autoNumberConfigRecord["trial_number"] : 0;
                        int seed = autoNumberConfigRecord.Contains("trial_seedvalue") ? (int)autoNumberConfigRecord["trial_seedvalue"] : 1;

                        // Create a new unique ID
                        string newId = $"{prefix}-{number}-{postfix}";
                        _traceService.Trace("new_id: " + newId);

                        // Update the unique ID field in the target record
                        Entity updateRecord = new Entity(entity, new Guid(idValue))
                        {
                            [field] = newId
                        };
                        _service.Update(updateRecord);
                        _traceService.Trace($"ID updated for {entity} record with ID {idValue}");

                        // Increment the number by the seed value
                        number += seed;
                        _traceService.Trace("next number: " + number);

                        // Update the next number in the auto number configuration record
                        Entity updateAutoNumberConfigRecord = new Entity(autoNumberConfigRecord.LogicalName, autoNumberConfigRecord.Id)
                        {
                            ["trial_number"] = number
                        };
                        _service.Update(updateAutoNumberConfigRecord);
                        _traceService.Trace($"Next number updated in {entity} auto number config record.");
                    }
                }
                else
                {
                    _traceService.Trace($"{entity} auto number config not present");
                }

                _traceService.Trace("End ProMX_AutoNumber Action");
            }
            catch (Exception ex)
            {
                // Log the exception details
                _traceService.Trace("Exception in ProMX_AutoNumber Execute Method: " + ex.Message);
                throw;
            }
        }
    }
}
