using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Base.Plugin
{
	public class EntityName : IPlugin
	{
		public void Execute(IServiceProvider serviceProvider)
		{
			// functionName to help debug at a glance
			string functionName = "Base.EntityName";
			IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

			// TODO - If you require tracing, uncomment the following line
			//ITracingService trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

			Entity entity = null;

			// Check if the InputParameters property bag contains a target
			// of the current operation and that target is of type DynamicEntity.
			if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
			{
				// Obtain the target business entity from the input parmameters.
				entity = (Entity)context.InputParameters["Target"];

				// Test for an entity type and message supported by your plug-in.
				if (context.PrimaryEntityName != Constants.EntityName.This) { return; }
				//if (context.MessageName != "<message>") { return; }

			}
			else
			{
				return;
			}
			try
			{
				IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
				IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

				Entity postImage = null;

				// Get the image we requested
				if (context.PostEntityImages.Contains(Constants.EntityName.This) && context.PostEntityImages[Constants.EntityName.This] is Entity)
				{
					postImage = (Entity)context.PostEntityImages[Constants.EntityName.This];
				}

				if (postImage == null)
					postImage = entity; // debugging

				// Only runs on create
				if (context.MessageName.ToLower() == "create")
				{
					// Magic goes here

					// BE SURE TO SIGN BEFORE DEPLOYING TO CRM (pfx key needed)
				}
			}

			catch (Exception ex)
			{
				string errorMsg = string.Format("An general Exception occurred in the {0} function: {1}", functionName, ex.Message);
				throw new InvalidPluginExecutionException(errorMsg, ex);
			}
		}
	}
}
