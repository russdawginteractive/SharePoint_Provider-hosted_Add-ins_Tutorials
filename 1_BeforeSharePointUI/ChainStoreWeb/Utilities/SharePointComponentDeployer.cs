// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using System;
using System.Data;
using System.Web;
using System.Linq;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using System.Data.SqlClient;


namespace ChainStoreWeb.Utilities
{
    public static class SharePointComponentDeployer
    {
        internal static SharePointContext sPContext;
        internal static Version localVersion;

        internal static Version RemoteTenantVersion
        {
            get
            {
                return GetTenantVersion();
            }
            set
            {
                SetTenantVersion(value);
            }
        }

        public static bool IsDeployed
        {
            get
            {
                if (RemoteTenantVersion < localVersion)
                    return false;
                else
                    return true;
            }
        }
        private static Version GetTenantVersion()
        {
            using (SqlConnection conn = SQLAzureUtilities.GetActiveSqlConnection())
            using (SqlCommand cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = "GetTenantVersion";
                cmd.CommandType = CommandType.StoredProcedure;
                SqlParameter name = cmd.Parameters.Add("@Name", SqlDbType.NVarChar);
                name.Value = sPContext.SPHostUrl.ToString();

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        reader.Read();
                        return new Version(reader["Version"].ToString());
                    }
                    else
                        throw new Exception("Unknown tenant: " + sPContext.SPHostUrl.ToString());
                }
            }//dispose conn and cmd
        }

        private static void SetTenantVersion(Version newVersion)
        {
            using (SqlConnection conn = SQLAzureUtilities.GetActiveSqlConnection())
            using (SqlCommand cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = "UpdateTenantVersion";
                cmd.CommandType = CommandType.StoredProcedure;
                SqlParameter name = cmd.Parameters.Add("@Name", SqlDbType.NVarChar);
                name.Value = sPContext.SPHostUrl.ToString();
                SqlParameter version = cmd.Parameters.Add("@Version", SqlDbType.NVarChar);
                version.Value = newVersion.ToString();
                cmd.ExecuteNonQuery();
            }//dispose conn and cmd
        }

        private static void CreateLocalEmployeesList()
        {
            using (var clientContext = sPContext.CreateUserClientContextForSPHost())
            {
                //Code for checking for existence of List with inefficient multiple calls to clientContext.LoadQuery
                var query = clientContext.Web.Lists.Where(x => x.Title == "Expected Shipments");
              
                IEnumerable<List> matchingLists = clientContext.LoadQuery(query);
                clientContext.ExecuteQuery();
                if (matchingLists.Count() != 0)
                {
                    // Create the list 
                    ListCreationInformation listInfo = new ListCreationInformation();
                    listInfo.Title = "Local Employees";
                    listInfo.TemplateType = (int)ListTemplateType.GenericList;
                    listInfo.Url = "Lists/Local Employees";

                    List localEmployeesList = clientContext.Web.Lists.Add(listInfo);
 

                    // Rename the Title field on the list 
                    Field field = localEmployeesList.Fields.GetByInternalNameOrTitle("Title");
                    field.Title = "Name";
  
                    field.Update();

                    // Add "Added to Corporate DB" field to the list 
                    localEmployeesList.Fields.AddFieldAsXml("<Field DisplayName='Added to Corporate DB'"
                        + " Type='Boolean'" 
                        + " ShowInEditForm='FALSE'"
                        + " ShowInNewForm='FALSE'>"
                        + "<Default>FALSE</Default></Field>",
                        true, AddFieldOptions.DefaultValue);

                    clientContext.ExecuteQuery();
                }
            }
          
        }

        private static void ChangeCustomActionRegistration()
        {
            using (var clientContext = sPContext.CreateUserClientContextForSPHost())
            {
                var query = clientContext.Web.UserCustomActions.Where(x => x.Name == "dc747915-c02a-4b66-a1c7-ead30fca94c6.AddEmployeeToCorpDB");
                IEnumerable<UserCustomAction> matchingActions = clientContext.LoadQuery(query);
                clientContext.ExecuteQuery();

                UserCustomAction webScopedEmployeeAction = matchingActions.Single();

                // Get a reference to the(empty) collection of custom actions
                // that are registered with the custom list.
                var queryForList = clientContext.Web.Lists.Where(x => x.Title == "Local Employees");
                IEnumerable<List> matchingLists = clientContext.LoadQuery(queryForList);
                clientContext.ExecuteQuery();

                List employeeList = matchingLists.First();
                var listActions = employeeList.UserCustomActions;
                clientContext.Load(listActions);
                listActions.Clear();

                // Add a blank custom action to the list's collection.
                var listScopedEmployeeAction = listActions.Add();

                // Copy property values from the descriptively deployed
                // custom action to the new custom action
                listScopedEmployeeAction.Title = webScopedEmployeeAction.Title;
                listScopedEmployeeAction.Location = webScopedEmployeeAction.Location;
                listScopedEmployeeAction.Sequence = webScopedEmployeeAction.Sequence;
                listScopedEmployeeAction.CommandUIExtension = webScopedEmployeeAction.CommandUIExtension;
                listScopedEmployeeAction.Update();

                // Delete the original custom action.         
                webScopedEmployeeAction.DeleteObject();
                // NOTE: Originally this action failed due to permissions.
                // Upon research of this issue, I found this Stack Overflow Article.
                // https://stackoverflow.com/questions/29675567/access-denied-office-365-sharepoint-online-with-global-admin-account
                // Based on the provided solution, I wrote a quick app to check those permissions
                // and indeed they were set to false as indicated. I reset them in the SP Admin
                // however, it takes 24 hours for that setting to initiate. Once 24 hours had passed 
                // I retried again and it was successful.
                /*
                 * This of course poses an undesirable solution due to having to reconfigure SP Admin to make something like this work.
                 * TODO: If this is found to be an issue in the future, This provided solution needs to be tried!
                 *  I will need to figure out how this actually works to get it to do this properly.
                 *  For the meantime, I will leave the SP Admin setting alone because it is only a Developer Account environment 
                 *  but this solution needs to be looked into if this is ever needed in the future.
Since any change to the scripting setting made through the SharePoint Online admin center may take up to 24 hours to take effect, you could enable scripting on a particular site collection immediately via CSOM API (SharePoint Online Client Components SDK) as demonstrated below:

public static void DisableDenyAddAndCustomizePages(ClientContext ctx, string siteUrl)
{
    var tenant = new Tenant(ctx);
    var siteProperties = tenant.GetSitePropertiesByUrl(siteUrl, true);
    ctx.Load(siteProperties);
    ctx.ExecuteQuery();

    siteProperties.DenyAddAndCustomizePages = DenyAddAndCustomizePagesStatus.Disabled;
    var result = siteProperties.Update();
    ctx.Load(result);
    ctx.ExecuteQuery();
    while (!result.IsComplete)
    {
        Thread.Sleep(result.PollingInterval);
        ctx.Load(result);
        ctx.ExecuteQuery();
    }
}
Usage

using (var ctx = GetContext(webUri, userName, password))
{
    using (var tenantAdminCtx = GetContext(tenantAdminUri, userName, password))
    {                  
         DisableDenyAddAndCustomizePages(tenantAdminCtx,webUri.ToString());
    }
    RegisterJQueryLibrary(ctx);
 }
where

public static void RegisterJQueryLibrary(ClientContext context)
{
    var actions = context.Site.UserCustomActions;
    var action = actions.Add();
    action.Location = "ScriptLink";
    action.ScriptSrc = "~SiteCollection/Style Library/Scripts/jQuery/jquery.min.js";
    action.Sequence = 1482;
    action.Update();
    context.ExecuteQuery();
}
                 * END OF TODO
                 */

                clientContext.ExecuteQuery();

            }
        }
        internal static void DeployChainStoreComponentsToHostWeb(HttpRequest request)
        {
            // Deployment code goes here.
            CreateLocalEmployeesList();
            ChangeCustomActionRegistration();
            RemoteTenantVersion = localVersion;
        }
    }
}

/*

OfficeDev/SharePoint_Provider-hosted_Add-ins_Tutorials, https://github.com/OfficeDev/SharePoint_Provider-hosted_Add-ins_Tutorials
 
Copyright (c) Microsoft Corporation
All rights reserved. 
 
MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
 
The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.    
  
*/
