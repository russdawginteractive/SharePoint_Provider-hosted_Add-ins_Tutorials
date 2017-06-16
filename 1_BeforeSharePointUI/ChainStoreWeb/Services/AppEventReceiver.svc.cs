using System;
using System.Data.SqlClient;
using System.Data;
using ChainStoreWeb.Utilities;
using Microsoft.SharePoint.Client.EventReceivers;
using Microsoft.SharePoint.Client;
using System.ServiceModel.Channels;
using System.ServiceModel;
using System.Net;

namespace ChainStoreWeb.Services
{
    public class AppEventReceiver : IRemoteEventService
    {
        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            // GIVES URL of Service Bus
            string debugEndpoint = System.ServiceModel.OperationContext.Current.Channel.LocalAddress.Uri.ToString();
            SPRemoteEventResult result = new SPRemoteEventResult();
            string tenantName = properties.AppEventProperties.HostWebFullUrl.ToString();
            if (!tenantName.EndsWith("/"))
            {
                tenantName += "/";
            }
            //using (ClientContext clientContext = TokenHelper.CreateAppEventClientContext(properties, useAppWeb: false))
            //{
            //    if (clientContext != null)
            //    {
            //        clientContext.Load(clientContext.Web);
            //        clientContext.ExecuteQuery();
            //    }
            //}

            switch (properties.EventType)
            {
                case SPRemoteEventType.AppInstalled:
                    // Custom installation logig goes here.
                    try
                    {
                        CreateTenant(tenantName);
                    }
                    catch (Exception e)
                    {
                        result.ErrorMessage = e.Message;
                        result.Status = SPRemoteEventServiceStatus.CancelWithError;
                    }
                    break;
                case SPRemoteEventType.AppUpgraded:
                    // This sample does not implemnet an add-in upgrade handler.
                    break;
                case SPRemoteEventType.AppUninstalling:
                    // Custom uninstallation logic goes here.
                    try
                    {
                        DeleteTenant(tenantName);
                    }
                    catch (Exception e)
                    {
                        // Tell SharePoint to cancel and roll back the event.
                        result.ErrorMessage = e.Message;
                        result.Status = SPRemoteEventServiceStatus.CancelWithError;
                    }
                    break;
            }
            // When a "before" event occurs (such as ItemAdding), call the event 
            // receiver code.
            //ListRemoteEventReceiver(properties);
            //return new SPRemoteEventResult();
            return result;
        }

        private void CreateTenant(string tenantName)
        {
            // Do not catch exceptions. Allow them to bubble up and trigger roll back
            // of installation.

            using (SqlConnection conn = SQLAzureUtilities.GetActiveSqlConnection())
            using (SqlCommand cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = "AddTenant";
                cmd.CommandType = CommandType.StoredProcedure;
                SqlParameter name = cmd.Parameters.Add("@Name", SqlDbType.NVarChar);
                name.Value = tenantName;
                cmd.ExecuteNonQuery();
            }//dispose conn and cmd
        }

        private void DeleteTenant(string tenantName)
        {
            // Do not catch exceptions. Allow them to bubble up and trigger roll back
            // on un-installation (removal from 2nd level Recycle Bin).

            using (SqlConnection conn = SQLAzureUtilities.GetActiveSqlConnection())
            using (SqlCommand cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = "RemoveTenant";
                cmd.CommandType = CommandType.StoredProcedure;
                SqlParameter name = cmd.Parameters.Add("@Name", SqlDbType.NVarChar);
                name.Value = tenantName;
                cmd.ExecuteNonQuery();
            } // dispose conn and cmd
        }
        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }
        public static void ListRemoteEventReceiver(SPRemoteEventProperties properties)
        {
            string logListTitle = "EventLog";

            // Return if the event is from the EventLog list itself. Otherwise, it may go into
            // an infinite loop.
            if (string.Equals(properties.ItemEventProperties.ListTitle, logListTitle,
                  StringComparison.OrdinalIgnoreCase))
                return;

            // Get the token from the request header.
            HttpRequestMessageProperty requestProperty =
                  (HttpRequestMessageProperty)OperationContext
                   .Current.IncomingMessageProperties[HttpRequestMessageProperty.Name];
            string contextTokenString = requestProperty.Headers["X-SP-ContextToken"];

            // If there is a valid token, continue.
            if (contextTokenString != null)
            {
                SharePointContextToken contextToken =
                    TokenHelper.ReadAndValidateContextToken(contextTokenString,
                         requestProperty.Headers[HttpRequestHeader.Host]);

                Uri sharepointUrl = new Uri(properties.ItemEventProperties.WebUrl);
                string accessToken = TokenHelper.GetAccessToken(contextToken,
                                                      sharepointUrl.Authority).AccessToken;
                bool exists = false;

                // Retrieve the log list "EventLog" and add the name of the event that occurred
                // to it with a date/time stamp.
                using (ClientContext clientContext =
                     TokenHelper.GetClientContextWithAccessToken(sharepointUrl.ToString(),
                                                                                                         accessToken))
                {
                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();
                    List logList = clientContext.Web.Lists.GetByTitle(logListTitle);

                    try
                    {
                        clientContext.Load(logList);
                        clientContext.ExecuteQuery();
                        exists = true;
                    }

                    catch (Microsoft.SharePoint.Client.ServerUnauthorizedAccessException)
                    {
                        // If the user doesn't have permissions to access the server that's 
                        // running SharePoint, return.
                        return;
                    }

                    catch (Microsoft.SharePoint.Client.ServerException)
                    {
                        // If an error occurs on the server that's running SharePoint, return.
                        exists = false;
                    }

                    // Create a log list called "EventLog" if it doesn't already exist.
                    if (!exists)
                    {
                        ListCreationInformation listInfo = new ListCreationInformation();
                        listInfo.Title = logListTitle;
                        // Create a generic custom list.
                        listInfo.TemplateType = 100;
                        clientContext.Web.Lists.Add(listInfo);
                        clientContext.Web.Context.ExecuteQuery();
                    }

                    // Add the event entry to the EventLog list.
                    string itemTitle = "Event: " + properties.EventType.ToString() +
                          " occurred on: " +
                          DateTime.Now.ToString(" yyyy/MM/dd/HH:mm:ss:fffffff");
                    ListCollection lists = clientContext.Web.Lists;
                    List selectedList = lists.GetByTitle(logListTitle);
                    clientContext.Load<ListCollection>(lists);
                    clientContext.Load<List>(selectedList);
                    ListItemCreationInformation listItemCreationInfo =
                          new ListItemCreationInformation();
                    var listItem = selectedList.AddItem(listItemCreationInfo);
                    listItem["Title"] = itemTitle;
                    listItem.Update();
                    clientContext.ExecuteQuery();
                }
            }
        }
    }
}
