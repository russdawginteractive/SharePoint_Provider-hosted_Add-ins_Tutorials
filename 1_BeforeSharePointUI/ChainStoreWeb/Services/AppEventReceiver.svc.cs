using System;
using System.Data.SqlClient;
using System.Data;
using ChainStoreWeb.Utilities;
using Microsoft.SharePoint.Client.EventReceivers;

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

    }
}
