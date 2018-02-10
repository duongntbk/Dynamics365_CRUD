using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Dynamics365_CRUD
{
    public class Dynamics365_CRUD : IPlugin
    {
        private IOrganizationService service;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Get service from service Provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);

            // Demo for Delete operation.
            demoDelete();

            // Demo for Update operation.
            demoUpdate();

            // Demo for Retrieve operation.
            var retrievedEntity = demoRetrieve();

            // Demo for Create operation.
            demoCreate(retrievedEntity);
        }

        /// <summary>
        /// Demo for Retrieve operation.
        /// Retrieve data from record with Employee Code equals NIBCREATE.
        /// </summary>
        /// <returns></returns>
        private Entity demoRetrieve()
        {
            var query = new QueryExpression()
            {
                // Physical name of entity we want to retrieve.
                EntityName = "new_employee",
                // All fields we want to retrieve.
                // If we do not specified ColumnSet, only Guid will be retrieved.
                // If we want to retrieve all columns, set ColumnSet = new ColumnSet(true).
                // Notice that if a field in record is not set (empty)
                // then the corresponding key will not exist in result array instead of being null.
                ColumnSet = new ColumnSet("new_employee_name", "new_dob", "new_gender", "new_employee_type", "new_manager"),
                // The filtering condition
                Criteria =
                {
                    // We actually do not need to specified FilterExpression here because the default value is LogicalOperator.And.
                    // Only retrieve records satisfy all condition in Conditions array.
                    FilterOperator = LogicalOperator.And,
                    // Look for record with Employee Code equals NIB00003 and is active.
                    Conditions =
                    {
                        new ConditionExpression("new_employee_code", ConditionOperator.Equal, "NIB00003"),
                        new ConditionExpression("statuscode", ConditionOperator.Equal, 1)
                    }
                }
            };

            // All results will be stored in an array called Entities, because we know there is at most 1 record with code equals NIB00003
            // we call FirstOrDefault here.
            var retrievedEntityList = service.RetrieveMultiple(query);
            return retrievedEntityList.Entities.FirstOrDefault();
        }

        /// <summary>
        /// Demo for Create operation.
        /// Create a new record with data from retrievedEntity.
        /// </summary>
        /// <param name="retrievedEntity"></param>
        private void demoCreate(Entity retrievedEntity)
        {
            // Retrieve the value of Employee Name, Date of Birth...
            var employeeName = getField<string>(retrievedEntity, "new_employee_name");
            var dob = getField<DateTime?>(retrievedEntity, "new_dob");
            var gender = getField<bool?>(retrievedEntity, "new_gender");
            var employeeType = getField<OptionSetValue>(retrievedEntity, "new_employee_type");
            var manager = getField<EntityReference>(retrievedEntity, "new_manager");

            // We will call Create method on this new entity object.
            // Setting value of Employee Name, Date of Birth...
            var newEntity = new Entity("new_employee");
            newEntity["new_employee_code"] = "NIBCREATE"; // Indicate that this is a record created using plugins.
            newEntity["new_employee_name"] = employeeName + "_" +  DateTime.Now.ToString("yyyyMMdd HH:mm:ss"); // Add a time stamp.
            newEntity["new_dob"] = dob;
            newEntity["new_gender"] = gender;
            newEntity["new_employee_type"] = employeeType;
            newEntity["new_manager"] = manager;

            // Create new record using Create method.
            service.Create(newEntity);
        }

        /// <summary>
        /// Demo for delete operation.
        /// Delete all records with Employee Code equals NIBCREATE, except the newest.
        /// </summary>
        private void demoDelete()
        {
            var query = new QueryExpression()
            {
                // Physical name of entity we want to retrieve.
                EntityName = "new_employee",
                // The filtering condition
                Criteria =
                {
                    // We actually do not need to specified FilterExpression here because the default value is LogicalOperator.And.
                    // Only retrieve records satisfy all condition in Conditions array.
                    FilterOperator = LogicalOperator.And,
                    // Look for record with Employee Code equals NIB00003 and is active.
                    Conditions =
                    {
                        new ConditionExpression("new_employee_code", ConditionOperator.Equal, "NIBCREATE"),
                        new ConditionExpression("statuscode", ConditionOperator.Equal, 1)
                    }
                },
                // Order results by create date in Descending order (newer record on top).
                Orders =
                {
                    new OrderExpression("createdon", OrderType.Descending)
                }
            };

            // All results will be stored in an array called Entities, because we know there is at most 1 record with code equals NIB00003
            // we call FirstOrDefault here.
            var retrievedEntityList = service.RetrieveMultiple(query);

            // If there is only 1 record with code equals NIBCREATE, we exit method here.
            // We do not want to delete the newest record with code equals NIBCREATE.
            if (retrievedEntityList.Entities.Count <= 1)
            {
                return;
            }
            
            // The record is the newest one, we do not want to delete it.
            for (var i = 1; i < retrievedEntityList.Entities.Count; i++)
            {
                service.Delete("new_employee", retrievedEntityList.Entities[i].Id);
            }
        }

        /// <summary>
        /// Demo for update operation.
        /// Change the Employee Name of newest record with code equals NIBCREATE to Nguyen Van Demo.
        /// </summary>
        private void demoUpdate()
        {
            var query = new QueryExpression()
            {
                // Physical name of entity we want to retrieve.
                EntityName = "new_employee",
                // The filtering condition
                Criteria =
                {
                    // We actually do not need to specified FilterExpression here because the default value is LogicalOperator.And.
                    // Only retrieve records satisfy all condition in Conditions array.
                    FilterOperator = LogicalOperator.And,
                    // Look for record with Employee Code equals NIB00003 and is active.
                    Conditions =
                    {
                        new ConditionExpression("new_employee_code", ConditionOperator.Equal, "NIBCREATE"),
                        new ConditionExpression("statuscode", ConditionOperator.Equal, 1)
                    }
                },
                // Order results by create date in Descending order (newer record on top).
                Orders =
                {
                    new OrderExpression("createdon", OrderType.Descending)
                }
            };

            // All results will be stored in an array called Entities, because we know there is at most 1 record with code equals NIB00003
            // we call FirstOrDefault here.
            var retrievedEntityList = service.RetrieveMultiple(query);

            // If we cannot find any record with code equals NIBCREATE, exit function.
            if (retrievedEntityList.Entities.Count < 1)
            {
                return;
            }

            // Create a new entity object with Guid of the first record in EntityCollection.
            // Always create a new entity instead of reuse entity in EntityCollection
            // because we don't want to update unwanted field.
            var updateEntity = new Entity("new_employee", retrievedEntityList.Entities.First().Id);
            // Set new Employee Name here.
            updateEntity["new_employee_name"] = "Nguyen Van Demo";
            // Call update function to update record.
            service.Update(updateEntity);
        }

        /// <summary>
        /// Retrive value of attribute fieldName from Entity entity if existed.
        /// If there is no such attribute, return default value of type T.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="entity"></param>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        private T getField<T>(Entity entity, string fieldName)
        {
            // Check if entity is null, entity does not contain fieldName, or if fieldName is null.
            if (entity == null || !entity.Contains(fieldName) || entity[fieldName] == null)
            {
                // Return defaul value of type T if cannot retrieve value of fieldName.
                return default(T);
            }

            // Cast fieldName to type T and return.
            return (T)entity[fieldName];
        } 
    }
}
