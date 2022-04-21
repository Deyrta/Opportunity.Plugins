using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace Opportunity.Plugins
{
    public class SetActivitiesRegardingFixed : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));


            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                Entity entity = (Entity)context.InputParameters["Target"];

                if (entity == null || !entity.Attributes.Contains("originatingleadid"))
                    return;

                EntityReference attributeValue = entity.GetAttributeValue<EntityReference>("originatingleadid");
                if(attributeValue != null)
                {
                    foreach(Entity relatedPhoneCall in this.GetRelatedPhoneCalls(attributeValue.Id, service))
                    {
                        // We cant set regarding to a completed activity with non admin user
                        if (relatedPhoneCall.GetAttributeValue<OptionSetValue>("statuscode").Value != 2 
                            && relatedPhoneCall.GetAttributeValue<OptionSetValue>("statuscode").Value != 3
                            && relatedPhoneCall.GetAttributeValue<OptionSetValue>("statuscode").Value != 4)
                        {
                            Entity entityReference = new Entity(relatedPhoneCall.LogicalName, relatedPhoneCall.Id);
                            entityReference["regardingobjectid"] = new EntityReference(entity.LogicalName, entity.Id);
                            service.Update(entityReference);
                        }
                    }
                }
            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("An error occurred in SetActivitiesRegardingFixed plugin.", ex);
            }

            catch (Exception ex)
            {
                tracingService.Trace("SetActivitiesRegardingFixed plugin: {0}", ex.ToString());
                throw;
            }
        }

        private List<Entity> GetRelatedPhoneCalls(Guid regardingObjId, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression()
            {
                ColumnSet = new ColumnSet("statuscode"),
                NoLock = true,
                EntityName = "phonecall"
            };
            FilterExpression filterExpression = new FilterExpression()
            {
                FilterOperator = LogicalOperator.And
            };
            filterExpression.Conditions.Add(new ConditionExpression("regardingobjectid", ConditionOperator.Equal, (object)regardingObjId));
            query.Criteria = filterExpression;
            return service.RetrieveMultiple(query).Entities.ToList<Entity>();
        }
    }
}
