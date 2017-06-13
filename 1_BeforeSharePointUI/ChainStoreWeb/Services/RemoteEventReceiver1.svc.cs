using ChainStoreWeb.Utilities;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace ChainStoreWeb.Services
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "RemoteEventReceiver1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select RemoteEventReceiver1.svc or RemoteEventReceiver1.svc.cs at the Solution Explorer and start debugging.
    public class RemoteEventReceiver1 : IRemoteEventService
    {
        /// <summary>
        /// Handles events that occur before an action occurs, 
        /// such as when a user is adding or deleting a list item.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        /// <returns>Holds information returned from the remote event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Handles events that occur after an action occurs, 
        /// such as after a user adds an item to a list or deletes an item from a list.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            switch (properties.EventType)
            {
                case SPRemoteEventType.ItemUpdated:
                    // Handle the item updated event.
                    switch (properties.ItemEventProperties.ListTitle)
                    {
                        case "Expected Shipments":
                            // Handle the arrival of a shipment.
                            bool updateComplete = TryUpdateInventory(properties);
                            if (updateComplete)
                            {
                                RecordInventoryUpdateLocally(properties);
                            }
                            break;
                    }
                    break;
            }
        }

        private void RecordInventoryUpdateLocally(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                List expectedShipmentsList = clientContext.Web.Lists.GetByTitle(properties.ItemEventProperties.ListTitle);
                ListItem arrivedItem = expectedShipmentsList.GetItemById(properties.ItemEventProperties.ListItemId);
                arrivedItem["Added_x0020_t0_x0020_Inventory"] = true;
                arrivedItem.Update();
                clientContext.ExecuteQuery();
            }
        }

        private bool TryUpdateInventory(SPRemoteEventProperties properties)
        {
            bool successFlag = false;

            // Test whether the list item is changing because the product has arrived
            // or for some other reason. If the former, add it to the inventory and set the success flag
            // to true.  
            try
            {
                // THIS PART THROWS AN ERROR UNLESS BOTH "Arrived" AND "Added to Inventory" have new values.
                var arrived = Convert.ToBoolean(properties.ItemEventProperties.AfterProperties["Arrived"]);
                var addedToInventory = Convert.ToBoolean(properties.ItemEventProperties.AfterProperties["Added_x0020_to_x0020_Inventory"]);

                if (arrived && !addedToInventory)
                {
                    // Add the item to inventory
                    // THIS PART DOES NOT WORK UNLESS "Title" and "Quantiy" have changed and have been sent.
                    using (SqlConnection conn = SQLAzureUtilities.GetActiveSqlConnection())
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        conn.Open();
                        cmd.CommandText = "UpdateInventory";
                        cmd.CommandType = CommandType.StoredProcedure;
                        SqlParameter tenant = cmd.Parameters.Add("@Tenant", SqlDbType.NVarChar);
                        tenant.Value = properties.ItemEventProperties.WebUrl + "/";
                        SqlParameter product = cmd.Parameters.Add("@ItemName", SqlDbType.NVarChar, 50);
                        product.Value = properties.ItemEventProperties.AfterProperties["Title"]; // not "Product"
                        SqlParameter quantity = cmd.Parameters.Add("@Quantity", SqlDbType.SmallInt);
                        quantity.Value = Convert.ToUInt16(properties.ItemEventProperties.AfterProperties["Quantity"]);
                        cmd.ExecuteNonQuery();
                    }

                    successFlag = true;
                }
            }
            catch (KeyNotFoundException)
            {
                successFlag = false;
            }
           
            return successFlag;
        }
    }
}
