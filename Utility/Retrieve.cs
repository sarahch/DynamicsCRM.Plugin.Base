using System;
using Microsoft.Xrm.Sdk;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Metadata;

namespace Base.Utility
{
	public static class Retrieve
	{
		/// <summary>
		/// The function gets the attrubute metadata specifically to translate between text and value for option sets
		/// </summary>
		/// <param name="entity">Name of the entity (Constants.ENTITYNAME.This)</param>
		/// <param name="service"></param>
		/// <param name="property"></param>
		/// <returns></returns>
		static public RetrieveAttributeResponse RetrieveAttributeMetaData(this Entity entity, IOrganizationService service, string property)
		{
			RetrieveAttributeRequest attribReq = new RetrieveAttributeRequest();
			attribReq.EntityLogicalName = entity.LogicalName;
			attribReq.LogicalName = property;
			attribReq.RetrieveAsIfPublished = true;
			return (RetrieveAttributeResponse)service.Execute(attribReq);
		}

		/// <summary>
		/// Retrieve all entity records using a field and value
		/// </summary>
		/// <param name="service"></param>
		/// <param name="entityName">Name of the entity (Constants.ENTITYNAME.This)</param>
		/// <param name="entitySearchField">Name of the entity field to filter by (Constants.ENTITYNAME.FIELD)</param>
		/// <param name="entitySearchFieldValue">Value for the entity field to filter by</param>
		/// <param name="columnSet"></param>
		/// <param name="activeRecordsOnly">Optional - Selects only records with statecode of 0, which is usually active, or open records</param>
		/// <returns></returns>
		static public List<Entity> EntitiesList(IOrganizationService service, string entityName, string entitySearchField, object entitySearchFieldValue, ColumnSet columnSet, bool activeRecordsOnly = false)
		{
			RetrieveMultipleRequest getRequest = new RetrieveMultipleRequest();
			QueryExpression qex = new QueryExpression(entityName);
			if (columnSet == null)
				qex.ColumnSet = new ColumnSet(true); // They get all
			else
				qex.ColumnSet = columnSet; // Give them just what they asked for

			qex.Criteria = new FilterExpression();
			qex.Criteria.FilterOperator = LogicalOperator.And;
			if (entitySearchFieldValue != null)
				qex.Criteria.AddCondition(new ConditionExpression(entitySearchField, ConditionOperator.Equal, entitySearchFieldValue));
			if (activeRecordsOnly)
				qex.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

			getRequest.Query = qex;

			RetrieveMultipleResponse returnValues = (RetrieveMultipleResponse)service.Execute(getRequest);

			if (returnValues.EntityCollection.Entities != null)
			{
				return returnValues.EntityCollection.Entities.ToList();
			}

			return null;
		}

		/// <summary>
		/// Retrieve all entity records using up to three fields and values
		/// </summary>
		/// <param name="service"></param>
		/// <param name="entityName">Name of the entity (Constants.ENTITYNAME.This)</param>
		/// <param name="entitySearchField1">Name of the 1st entity field to filter by (Constants.ENTITYNAME.FIELD)</param>
		/// <param name="entitySearchFieldValue1">Value for the 1st entity field to filter by</param>
		/// <param name="entitySearchField2">Name of the 2nd entity field to filter by (Constants.ENTITYNAME.FIELD)</param>
		/// <param name="entitySearchFieldValue2">Value for the 2nd entity field to filter by</param>
		/// <param name="entitySearchField3">Name of the 3rd entity field to filter by (Constants.ENTITYNAME.FIELD)</param>
		/// <param name="entitySearchFieldValue3">Value for the 3rd entity field to filter by</param>
		/// <param name="columnSet"></param>
		/// <param name="activeRecordsOnly">Optional - Selects only records with statecode of 0, which is usually active, or open records</param>
		/// <returns></returns>
		static public List<Entity> EntitiesList(IOrganizationService service, string entityName, string entitySearchField1, object entitySearchFieldValue1,
			string entitySearchField2, object entitySearchFieldValue2, string entitySearchField3, object entitySearchFieldValue3, ColumnSet columnSet, bool activeRecordsOnly = false)
		{
			RetrieveMultipleRequest getRequest = new RetrieveMultipleRequest();
			QueryExpression qex = new QueryExpression(entityName);
			if (columnSet == null)
				qex.ColumnSet = new ColumnSet(true); // They get all
			else
				qex.ColumnSet = columnSet; // Give them just what they asked for

			qex.Criteria = new FilterExpression();
			qex.Criteria.FilterOperator = LogicalOperator.And;
			if (entitySearchFieldValue1 != null)
				qex.Criteria.AddCondition(new ConditionExpression(entitySearchField1, ConditionOperator.Equal, entitySearchFieldValue1));
			if (entitySearchFieldValue2 != null)
				qex.Criteria.AddCondition(new ConditionExpression(entitySearchField2, ConditionOperator.Equal, entitySearchFieldValue2));
			if (entitySearchFieldValue3 != null)
				qex.Criteria.AddCondition(new ConditionExpression(entitySearchField3, ConditionOperator.Equal, entitySearchFieldValue3));
			if (activeRecordsOnly)
				qex.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

			getRequest.Query = qex;

			RetrieveMultipleResponse returnValues = (RetrieveMultipleResponse)service.Execute(getRequest);

			if (returnValues.EntityCollection.Entities != null)
			{
				return returnValues.EntityCollection.Entities.ToList();
			}

			return null;
		}
		/// <summary>
		///
		/// </summary>
		/// <param name="oAttribute"></param>
		/// <param name="attMetadata"></param>
		/// <returns></returns>
		static public string GetMetadataValue(object oAttribute, RetrieveAttributeResponse attMetadata)
		{
			string sReturn = string.Empty;
			if (oAttribute.GetType().Equals(typeof(Microsoft.Xrm.Sdk.OptionSetValue)))
			{
				OptionMetadata[] optionList = null;

				if (attMetadata.AttributeMetadata.GetType().FullName.Contains("PicklistAttributeMetadata"))
				{
					PicklistAttributeMetadata retrievedPicklistAttributeMetadata =
						(PicklistAttributeMetadata)attMetadata.AttributeMetadata;
					// Get the current options list for the retrieved attribute.
					optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();
				}
				else if (attMetadata.AttributeMetadata.GetType().FullName.Contains("StatusAttributeMetadata"))
				{
					StatusAttributeMetadata retrievedPicklistAttributeMetadata =
						(StatusAttributeMetadata)attMetadata.AttributeMetadata;
					// Get the current options list for the retrieved attribute.
					optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();
				}
				else if (attMetadata.AttributeMetadata.GetType().FullName.Contains("StateAttributeMetadata"))
				{
					StateAttributeMetadata retrievedPicklistAttributeMetadata =
						(StateAttributeMetadata)attMetadata.AttributeMetadata;
					// Get the current options list for the retrieved attribute.
					optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();
				}
				else
					return string.Empty;

				// get the text values
				int i = int.Parse((oAttribute as Microsoft.Xrm.Sdk.OptionSetValue).Value.ToString());
				for (int c = 0; c < optionList.Length; c++)
				{
					OptionMetadata opmetadata = (OptionMetadata)optionList.GetValue(c);
					if (opmetadata.Value == i)
					{
						sReturn = opmetadata.Label.UserLocalizedLabel.Label;
						break;
					}
				}

			}
			else if (oAttribute.GetType().Equals(typeof(Microsoft.Xrm.Sdk.Money)))
			{
				sReturn = (oAttribute as Microsoft.Xrm.Sdk.Money).Value.ToString();
			}
			else if (oAttribute.GetType().Equals(typeof(Microsoft.Xrm.Sdk.EntityReference)))
			{
				sReturn = (oAttribute as Microsoft.Xrm.Sdk.EntityReference).Name;
			}
			else
			{
				sReturn = oAttribute.ToString();
			}
			if (sReturn == null || sReturn.Length == 0)
				sReturn = "No Value";
			return sReturn;
		}

		/// <summary>
		/// Get entity from entityref
		/// </summary>
		/// <param name="EntityRef">Name of the entity (Constants.ENTITYNAME.This)</param>
		/// <param name="service"></param>
		/// <param name="columnSet"></param>
		/// <returns></returns>
		static public Entity GetEntity(this EntityReference EntityRef, IOrganizationService service, ColumnSet columnSet)
		{
			return Retrieve.GetEntity(service, EntityRef.LogicalName, EntityRef.LogicalName + "id", EntityRef.Id, columnSet);
		}

		/// <summary>
		/// Get entity by the EntityId
		/// </summary>
		/// <param name="Entity">Name of the entity (Constants.ENTITYNAME.This)</param>
		/// <param name="service"></param>
		/// <param name="columnSet"></param>
		/// <returns></returns>
		static public Entity GetEntity(this Entity Entity, IOrganizationService service, ColumnSet columnSet)
		{
			return Retrieve.GetEntity(service, Entity.LogicalName, Entity.LogicalName + "id", Entity.Id, columnSet);
		}

		/// <summary>
		/// Get a single entity record based on a specific field and value, used when there may be more than one
		/// </summary>
		/// <param name="service"></param>
		/// <param name="entityName">Name of the entity (Constants.ENTITYNAME.This)</param>
		/// <param name="entityField">Name of the entity field to filter by (Constants.ENTITYNAME.FIELD)</param>
		/// <param name="entityFieldValue">Value for the entity field to filter by</param>
		/// <param name="columnSet"></param>
		/// <param name="activeRecordsOnly">Optional - Selects only records with statecode of 0, which is usually active, or open records</param>
		/// <returns></returns>
		static public Entity GetEntity(IOrganizationService service, string entityName, string entityField, object entityFieldValue, ColumnSet columnSet, bool activeRecordsOnly = false)
		{
			List<Entity> entities = EntitiesList(service, entityName, entityField, entityFieldValue, columnSet, activeRecordsOnly);

			if (entities != null && entities.Count() >= 1)
			{
				return (Entity)entities.First();
			}

			return null;
		}

		/// <summary>
		///  Get a single entity record based on two specific fields and values, used when there may be more than one
		/// </summary>
		/// <param name="service"></param>
		/// <param name="entityName">Name of the entity (Constants.ENTITYNAME.This)</param>
		/// <param name="entityField1">Name of the 1st entity field to filter by (Constants.ENTITYNAME.FIELD)</param>
		/// <param name="entityFieldValue1">Value for the 1st entity field to filter by</param>
		/// <param name="entityField2">Name of the 2nd entity field to filter by (Constants.ENTITYNAME.FIELD)</param>
		/// <param name="entityFieldValue2">Value for the 2nd entity field to filter by</param>
		/// <param name="columnSet"></param>
		/// <param name="activeRecordsOnly">Optional - Selects only records with statecode of 0, which is usually active, or open records</param>
		/// <returns></returns>
		static public Entity GetEntityBy2(IOrganizationService service, string entityName, string entityField1, object entityFieldValue1, string entityField2, object entityFieldValue2, ColumnSet columnSet, bool activeRecordsOnly = false)
		{
			List<Entity> entities = EntitiesList(service, entityName, entityField1, entityFieldValue1, entityField2, entityFieldValue2, null, null, columnSet, activeRecordsOnly);

			if (entities != null && entities.Count() >= 1)
			{
				return (Entity)entities.First();
			}

			return null;
		}


		/// <summary>
		/// If they don't give us the direct id of a related entity, we can still find it
		/// </summary>
		/// <param name="service"></param>
		/// <param name="entityName"></param>
		/// <param name="primaryIdName"></param>
		/// <param name="primaryId"></param>
		/// <param name="secondaryIdName"></param>
		/// <param name="secondaryId"></param>
		/// <param name="activeRecordsOnly">Optional - Selects only records with statecode of 0, which is usually active, or open records</param>
		/// <returns></returns>
		public static Guid RetrieveRelatedEntityId(IOrganizationService service, string entityName, string primaryIdName, Guid primaryId, string secondaryIdName, Guid secondaryId, bool activeRecordsOnly = false)
		{
			// FetchXML queries - faster and lighter when we can use them!
			string fetchXML = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
			fetchXML += "   <entity name='" + entityName + "'>";
			fetchXML += "       <attribute name='" + entityName + "id' />";
			fetchXML += "       <filter type='and'>";
			fetchXML += "         <condition attribute='" + primaryIdName + "' operator='eq' value='";
			fetchXML += primaryId;
			fetchXML += "' />";
			fetchXML += "         <condition attribute='" + secondaryIdName + "' operator='eq' value='";
			fetchXML += secondaryId;
			if (activeRecordsOnly)
				fetchXML += "         <condition attribute='statecode' operator='eq' value='0";
			fetchXML += "' />";
			fetchXML += "       </filter>";
			fetchXML += "     </entity>";
			fetchXML += "   </fetch>";

			EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXML));
			Guid id = Guid.Empty;
			if (result.Entities.Count > 0)
				id = result[0].Id;

			return id;
		}

		/// <summary>
		/// Gets an optionset value from an entity
		/// </summary>
		/// <param name="entity">Name of the entity (Constants.ENTITYNAME.This)</param>
		/// <param name="field">Name of the entity field </param>
		/// <param name="property">OptionSet</param>
		/// <param name="service"></param>
		/// <returns></returns>
		static public object getEntityValue(this Entity entity, string field, string property, IOrganizationService service)
		{
			if (entity != null && !entity.Contains(field))
				return null;

			ColumnSet colSet = new ColumnSet(property);

			Entity childEntity = Retrieve.GetEntity((EntityReference)entity[field], service, colSet);

			object propertyValue = null;

			if (childEntity != null && childEntity.Contains(property))
				propertyValue = childEntity[property];

			return propertyValue;
		}
	}
}
