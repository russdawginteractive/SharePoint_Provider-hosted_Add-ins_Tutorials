// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using System;
using System.Data.SqlClient;
using System.Data;
using ChainStoreWeb.Utilities;
using Microsoft.SharePoint.Client;
using System.Linq;
using System.Collections.Generic;

namespace ChainStoreWeb.Pages
{
    public partial class OrderForm : System.Web.UI.Page
    {
        protected SharePointContext spContext;

        protected void Page_Load(object sender, EventArgs e)
        {
            spContext = Session["SPContext"] as SharePointContext;
        }

        protected void btnCreateOrder_Click(object sender, EventArgs e)
        {
            UInt16 quantity;
            UInt16.TryParse(txtBoxQuantity.Text, out quantity);

            // Handle case where user presses the button without first entering rquired data.
            if (String.IsNullOrEmpty(txtBoxSupplier.Text) || String.IsNullOrEmpty(txtBoxItemName.Text))
            {
                lblOrderPrompt.Text = "Please enter a supplier and item.";
                lblOrderPrompt.ForeColor = System.Drawing.Color.Red;
                return;
            }
            else
            {
                if (quantity == 0)
                {
                    lblOrderPrompt.Text = "Quantity must be a positive number below 32,768.";
                    lblOrderPrompt.ForeColor = System.Drawing.Color.Red;
                    return;
                }
            }
            CreateExpectedShipment(txtBoxSupplier.Text, txtBoxItemName.Text, quantity);
            CreateOrder(txtBoxSupplier.Text, txtBoxItemName.Text, quantity);
        }

        private void CreateExpectedShipment(string supplier, string product, UInt16 quantity)
        {
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                // Original code.
                //List expectedShipmentsList = clientContext.Web.Lists.GetByTitle("Expected Shipments");
                //ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                //ListItem newItem = expectedShipmentsList.AddItem(itemCreateInfo);

                //newItem["Title"] = product;
                //newItem["Supplier"] = supplier;
                //newItem["Quantity"] = quantity;
                //newItem.Update();
                //clientContext.ExecuteQuery();
                // Code for checking for existence of List with inefficient multiple calls to clientContext.LoadQuery
                //var query = clientContext.Web.Lists.Where(x => x.Title == "Expected Shipments");
                ////query = from list in clientContext.Web.Lists
                ////            where list.Title == "Expected Shipments"
                ////            select list;
                //IEnumerable<List> matchingLists = clientContext.LoadQuery(query);
                //clientContext.ExecuteQuery();
                //if (matchingLists.Count() != 0)
                //{
                //    List expectedShipmentsList = matchingLists.Single();
                //    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                //    ListItem newItem = expectedShipmentsList.AddItem(itemCreateInfo);

                //    newItem["Title"] = product;
                //    newItem["Supplier"] = supplier;
                //    newItem["Quantity"] = quantity;
                //    newItem.Update();
                //}
                //clientContext.ExecuteQuery();

                // NOTE: Write code for checking for List via ConditionalScope.
                // https://msdn.microsoft.com/en-us/library/office/ee535891(v=office.14).aspx
                // What I found is that the code below does not work, it generates an Exception.
                // The Exception is:
                //      Incorrect usage of conditional scope.Some actions, such as setting a property or invoking a method, are not allowed inside a conditional scope.
                // The primary issue is that the ConditionalScope needs to be used ONLY for load operations.
                // For code like what is illustrated below, we need to use ExceptionScope. See Below...
                /* START OF EXCEPTION CODE
                List expectedShipmentsList = clientContext.Web.Lists.GetByTitle("Expected Shipments");
                
                ConditionalScope scope = new ConditionalScope(clientContext,
                    () => expectedShipmentsList.ServerObjectIsNull.Value != true);
                using (scope.StartScope())
                {
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem newItem = expectedShipmentsList.AddItem(itemCreateInfo);
                    newItem["Title"] = product;
                    newItem["Supplier"] = supplier;
                    newItem["Quantity"] = quantity;
                    newItem.Update();
                }
                
                clientContext.ExecuteQuery();
                lblResult.Text = scope.TestResult.Value.ToString();
                END OF EXCEPTION CODE */

                // NOTE: Begin ExceptionScope handling code to allow for setting of properties and invoking methods
                // http://ranaictiu-technicalblog.blogspot.com/2010/08/sharepoint-2010-exception-handling.html
                // https://msdn.microsoft.com/en-us/library/office/ee534976(v=office.14).aspx

                ExceptionHandlingScope scope = new ExceptionHandlingScope(clientContext);

                using (scope.StartScope())
                {
                    using (scope.StartTry())
                    {
                        List expectedShipmentsList = clientContext.Web.Lists.GetByTitle("Expected Shipments");
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newItem = expectedShipmentsList.AddItem(itemCreateInfo);

                        newItem["Title"] = product;
                        newItem["Supplier"] = supplier;
                        newItem["Quantity"] = quantity;
                        newItem.Update();
                    }
                    using (scope.StartCatch())
                    {
                        lblResult.Text = scope.ErrorMessage;
                    }
                    using (scope.StartFinally())
                    {
                        lblResult.Text = "Expected Shipments Updated Properly. Product Ordered!";
                    }
                }
                clientContext.ExecuteQuery();
            }
        }
        private void CreateOrder(String supplierName, String productName, UInt16 quantityOrdered)
        {
            using (SqlConnection conn = SQLAzureUtilities.GetActiveSqlConnection())
            using (SqlCommand cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = "AddOrder";
                cmd.CommandType = CommandType.StoredProcedure;
                SqlParameter tenant = cmd.Parameters.Add("@Tenant", SqlDbType.NVarChar);
                tenant.Value = spContext.SPHostUrl.ToString();
                SqlParameter supplier = cmd.Parameters.Add("@Supplier", SqlDbType.NVarChar, 50);
                supplier.Value = supplierName;
                SqlParameter itemName = cmd.Parameters.Add("@ItemName", SqlDbType.NVarChar, 50);
                itemName.Value = productName;
                SqlParameter quantity = cmd.Parameters.Add("@Quantity", SqlDbType.SmallInt);
                quantity.Value = quantityOrdered;
                cmd.ExecuteNonQuery();
            }
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