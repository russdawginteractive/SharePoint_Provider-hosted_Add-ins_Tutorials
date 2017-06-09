﻿// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

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

        internal static void DeployChainStoreComponentsToHostWeb(HttpRequest request)
        {
            // Deployment code goes here.
            CreateLocalEmployeesList();
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
