﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace SageIntegration
{
    public class SO_OrderBatch
    {
        public enum Status
        {
            Working = 0,
            Completed = 1
        }

        public string BatchNo = "";
        public string Username = "";
        public string Password = "";
        public string Company = "";
        public string BatchDescription = "";

        public static event BatchComplete Completed;
        public delegate void BatchComplete(List<SalesOrder> salesorders);

        public static event UpdateStatus StatusChanged;
        public delegate void UpdateStatus(Status s, int total, int current);

        public static event OrderPostedEvent OrderPosted;
        public delegate void OrderPostedEvent(SalesOrder salesorder);

        private List<SalesOrder> _salesorders = new List<SalesOrder>();
        private List<SalesOrder> _errorsalesorders = new List<SalesOrder>();

        public SO_OrderBatch(string Company = "", string Username = "", string Password = "", string BatchNo = "", string BatchDescription = "")
        {
            this.Company = Company;
            this.Username = Username;
            this.Password = Password;
            this.BatchNo = BatchNo;
            this.BatchDescription = BatchDescription;
        }

        public List<SalesOrder> SalesOrders()
        {
            return _salesorders;
        }
        public SalesOrder AddOrder(SalesOrder SalesOrder = null)
        {
            if (SalesOrder == null)
            {
                SalesOrder = new SalesOrder(); ;
            }
            _salesorders.Add(SalesOrder);

            return SalesOrder;
        }

        public void AddOrders(string[] salesorders)
        {
            foreach (string order in salesorders)
            {
                SalesOrder salesorder = new SalesOrder();
                salesorder.SalesOrderNo = order;
                _salesorders.Add(salesorder);
            }
        }

        public void MarkError(SalesOrder salesorder)
        {
            _salesorders.Remove(salesorder);
            _errorsalesorders.Add(salesorder);

        }

        public List<SalesOrder> Post()
        {

            // Grab init path from the registry
            Microsoft.Win32.RegistryKey SageKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\SOTAMAS90");
            string InitPath = SageKey.GetValue("Directory") + @"\Home";
            object retVal = new object();

            // Open Initial PVX object
            using (DispatchObject pvx = new DispatchObject("ProvideX.Script"))
            {
                // Set MAS90 Path
                pvx.InvokeMethod("Init", InitPath);

                // Set up my Session connection
                using (DispatchObject oSS = new DispatchObject(pvx.InvokeMethod("NewObject", "SY_Session")))
                {
                    // Set Username and password
                    // Set Company, and set the date
                    // Basic setup for the oSS object
                    oSS.InvokeMethod("nSetUser", Username, Password);
                    oSS.InvokeMethod("nLogon");
                    oSS.InvokeMethod("nSetCompany", Company);
                    oSS.InvokeMethod("nSetDate", "S/O", DateTime.Today.ToString("yyyyMMdd"));
                    oSS.InvokeMethod("nSetModule", "A/R");

                    // Set task ID to oSS
                    int TaskID = (int)oSS.InvokeMethod("nLookupTask", "SO_SalesOrder_ui");
                    retVal = oSS.InvokeMethod("nSetProgram", TaskID);

                    // SafeProcess just checks the last error message value and prints it to
                    // the debug output. In the future, I will have it post this to an error log
                    SafeProcess(oSS.InvokeMethod("nSetProgram", TaskID), oSS);

                    // Loop through the orders
                    // This saves time over opening the oSS and PVX object over and over again
                    foreach (SalesOrder Order in _salesorders)
                    {
                        // CurrentShipment.Packages.FindIndex(a => a == CurrentPackage)
                        if (StatusChanged != null)
                        {
                            StatusChanged(Status.Working, _salesorders.Count(), (_salesorders.FindIndex(a => a == Order) + 1));
                        }

                        // Open object for sales order
                        using (DispatchObject so_order = new DispatchObject(pvx.InvokeMethod("NewObject", "SO_SalesOrder_bus", oSS.GetObject())))
                        {
                            // Get the next salesorder number. 
                            object[] nextOrderNum = new object[] { "" };
                            SafeProcess(so_order.InvokeMethodByRef("nGetNextSalesOrderNo", nextOrderNum), so_order);
                            SafeProcess(so_order.InvokeMethod("nSetKey", nextOrderNum), so_order);
                            // Update the order objerct to the correct number
                            Order.SalesOrderNo = (string)nextOrderNum[0];

                            // Set Customer name and the ARDivision. We always use 00 as the ARDivision
                            SafeProcess(so_order.InvokeMethod("nSetValue", "ARDivisionNo$", "00"), so_order);
                            SafeProcess(so_order.InvokeMethod("nSetValue", "CustomerNo$", Order.CustomerNo), so_order);

                            SafeProcess(so_order.InvokeMethod("nSetValue", "OrderDate$", Order.OrderDate.ToString("yyyyMMdd")), so_order);
                            SafeProcess(so_order.InvokeMethod("nSetValue", "ShipExpireDate$", Order.RequiredDate.ToString("yyyyMMdd")), so_order);
                            SafeProcess(so_order.InvokeMethod("nSetValue", "ShipToZipcode$", Order.ShipToZipcode), so_order);
                            SafeProcess(so_order.InvokeMethod("nSetValue", "ShipToName$", Order.ShipToName), so_order);
                            SafeProcess(so_order.InvokeMethod("nSetValue", "ShipToAddress1$", Order.ShipToAddress1), so_order);
                            SafeProcess(so_order.InvokeMethod("nSetValue", "ShipToAddress2$", Order.ShipToAddress2), so_order);
                            SafeProcess(so_order.InvokeMethod("nSetValue", "CustomerPONo$", Order.CustomerPONo), so_order);
                            SafeProcess(so_order.InvokeMethod("nSetValue", "Comment$", Order.GiftHeader), so_order);

                            // Loop through each lineitem in the order
                            foreach (LineItem Line in Order.LineItems())
                            {
                                // Open object for Line item
                                using (DispatchObject so_line = new DispatchObject(so_order.GetProperty("oLines")))
                                {
                                    //SafeProcess(so_line.InvokeMethod("nAddLine"), so_line);
                                    retVal = so_line.InvokeMethod("nAddLine");
                                    // Set Item Code
                                    SafeProcess(so_line.InvokeMethod("nSetValue", "ItemCode$", Line.ItemCode), so_line);
                                    // Set Line Number
                                    //SafeProcess(so_line.InvokeMethod("nSetValue", "UDF_LINENO$", Line.LineNo), so_line);
                                    // Set Item Quantity
                                    SafeProcess(so_line.InvokeMethod("nSetValue", "QuantityOrdered", Line.Quantity), so_line);
                                    // Set Monogram
                                    SafeProcess(so_line.InvokeMethod("nSetValue", "UDF_PERSONALIZE1$", Line.Monogram), so_line);

                                    SafeProcess(so_line.InvokeMethod("nSetValue", "UDF_GIFTMESSAGE$", Order.GiftMessage), so_line);

                                    //Set LineKey
                                    object[] oLineKey = new object[] { "LineKey", "" };
                                    SafeProcess(so_line.InvokeMethodByRef("nGetValue", oLineKey), so_line);
                                    Line.LineKey = oLineKey[1].ToString();

                                    // Set Shipvia
                                    if (Order.Shipvia != string.Empty)
                                    {
                                        SafeProcess(so_order.InvokeMethod("nSetValue", "ShipVia$", Order.Shipvia), so_order);
                                    }
                                    // Set Descriptions
                                    if (Line.Description != string.Empty)
                                    {
                                        string Description = string.Empty;
                                        // Read Create an object to hold the response
                                        object[] getValueParam2 = new object[] { "ItemCodeDesc$", "" };

                                        // Read Into the object
                                        SafeProcess(so_line.InvokeMethodByRef("nGetValue", getValueParam2), so_line);

                                        // Retreive our description
                                        Description = getValueParam2[1].ToString();

                                        // Append the custom Description text and write the record
                                        Description = Description + Environment.NewLine + Line.Description;
                                        SafeProcess(so_line.InvokeMethod("nSetValue", "ItemCodeDesc$", Description), so_line);

                                    }

                                    retVal = so_line.InvokeMethod("nWrite");
                                    if ((int)retVal != 1)
                                    {
                                        // Line couldn't be added, or there was an issue with it. Handle the error message and continue
                                        // to the next line item
                                        continue;
                                    }

                                }
                            }

                            retVal = so_order.InvokeMethod("nWrite");
                            if ((int)retVal != 1) { Debug.WriteLine(so_order.GetProperty("sLastErrorMsg")); }
                            else
                            {
                                if (OrderPosted != null) { OrderPosted(Order); }

                                if (Order.ID != 0)
                                {

                                }
                            }
                        }

                    }

                    if (Completed != null) { Completed(_salesorders); }
                    if (StatusChanged != null) { StatusChanged(Status.Completed, 0, 0); }
                }

            }

            return _salesorders;
        }

        public void Sort()
        {
            this._salesorders = this._salesorders.OrderBy(c => c.LineItems().Count()).ThenBy(c => c.LineItems()[0].ItemCode).ToList();
        }

        public void Delete()
        {
            Microsoft.Win32.RegistryKey SageKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\SOTAMAS90");
            string InitPath = SageKey.GetValue("Directory") + @"\Home";
            object retVal;

            // Open Initial PVX object
            using (DispatchObject pvx = new DispatchObject("ProvideX.Script"))
            {
                // Set MAS90 Path
                pvx.InvokeMethod("Init", InitPath);

                // Set up my Session connection
                using (DispatchObject oSS = new DispatchObject(pvx.InvokeMethod("NewObject", "SY_Session")))
                {
                    // Set Username and password
                    // Set Company, and set the date
                    // Basic setup for the oSS object
                    oSS.InvokeMethod("nSetUser", Username, Password);
                    oSS.InvokeMethod("nLogon");
                    oSS.InvokeMethod("nSetCompany", Company);
                    oSS.InvokeMethod("nSetDate", "S/O", DateTime.Today.ToString("yyyyMMdd"));
                    oSS.InvokeMethod("nSetModule", "A/R");

                    // Set task ID to oSS
                    int TaskID = (int)oSS.InvokeMethod("nLookupTask", "SO_SalesOrder_ui");
                    retVal = oSS.InvokeMethod("nSetProgram", TaskID);

                    // SafeProcess just checks the last error message value and prints it to
                    // the debug output. In the future, I will have it post this to an error log
                    SafeProcess(oSS.InvokeMethod("nSetProgram", TaskID), oSS);

                    // Loop through the orders
                    // This saves time over opening the oSS and PVX object over and over again
                    foreach (SalesOrder salesorder in this.SalesOrders())
                    {
                        Debug.WriteLine(">>Changing " + salesorder.SalesOrderNo + "<<");
                        // Open object for sales order
                        using (DispatchObject so_order = new DispatchObject(pvx.InvokeMethod("NewObject", "SO_SalesOrder_bus", oSS.GetObject())))
                        {

                            SafeProcess(so_order.InvokeMethod("nSetKey", salesorder.SalesOrderNo), so_order);

                            SafeProcess(so_order.InvokeMethod("nDelete", salesorder.SalesOrderNo), so_order);
                        }

                        Debug.WriteLine("<<End Changing Sales Order<<");
                    }
                }

            }
        }

        public void ChangeDate(string date)
        {
            Microsoft.Win32.RegistryKey SageKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\SOTAMAS90");
            string InitPath = SageKey.GetValue("Directory") + @"\Home";
            object retVal;

            // Open Initial PVX object
            using (DispatchObject pvx = new DispatchObject("ProvideX.Script"))
            {
                // Set MAS90 Path
                pvx.InvokeMethod("Init", InitPath);

                // Set up my Session connection
                using (DispatchObject oSS = new DispatchObject(pvx.InvokeMethod("NewObject", "SY_Session")))
                {
                    // Set Username and password
                    // Set Company, and set the date
                    // Basic setup for the oSS object
                    oSS.InvokeMethod("nSetUser", Username, Password);
                    oSS.InvokeMethod("nLogon");
                    oSS.InvokeMethod("nSetCompany", Company);
                    oSS.InvokeMethod("nSetDate", "S/O", DateTime.Today.ToString("yyyyMMdd"));
                    oSS.InvokeMethod("nSetModule", "A/R");

                    // Set task ID to oSS
                    int TaskID = (int)oSS.InvokeMethod("nLookupTask", "SO_SalesOrder_ui");
                    retVal = oSS.InvokeMethod("nSetProgram", TaskID);

                    // SafeProcess just checks the last error message value and prints it to
                    // the debug output. In the future, I will have it post this to an error log
                    SafeProcess(oSS.InvokeMethod("nSetProgram", TaskID), oSS);

                    // Loop through the orders
                    // This saves time over opening the oSS and PVX object over and over again
                    foreach (SalesOrder salesorder in this.SalesOrders())
                    {
                        Debug.WriteLine(">>Changing " + salesorder.SalesOrderNo + "<<");
                        // Open object for sales order
                        using (DispatchObject so_order = new DispatchObject(pvx.InvokeMethod("NewObject", "SO_SalesOrder_bus", oSS.GetObject())))
                        {

                            SafeProcess(so_order.InvokeMethod("nSetKey", salesorder.SalesOrderNo), so_order);

                            SafeProcess(so_order.InvokeMethod("nSetValue", "ShipExpireDate$", date), so_order);

                            retVal = so_order.InvokeMethod("nWrite");
                            if ((int)retVal != 1)
                            {
                                Debug.WriteLine(so_order.GetProperty("sLastErrorMsg"));
                            }
                        }

                        Debug.WriteLine("<<End Changing Sales Order<<");
                    }
                }

            }
        }

        public void ChangeShipvia(string shipvia)
        {
            Microsoft.Win32.RegistryKey SageKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\SOTAMAS90");
            string InitPath = SageKey.GetValue("Directory") + @"\Home";
            object retVal;

            // Open Initial PVX object
            using (DispatchObject pvx = new DispatchObject("ProvideX.Script"))
            {
                // Set MAS90 Path
                pvx.InvokeMethod("Init", InitPath);

                // Set up my Session connection
                using (DispatchObject oSS = new DispatchObject(pvx.InvokeMethod("NewObject", "SY_Session")))
                {
                    // Set Username and password
                    // Set Company, and set the date
                    // Basic setup for the oSS object
                    oSS.InvokeMethod("nSetUser", Username, Password);
                    oSS.InvokeMethod("nLogon");
                    oSS.InvokeMethod("nSetCompany", Company);
                    oSS.InvokeMethod("nSetDate", "S/O", DateTime.Today.ToString("yyyyMMdd"));
                    oSS.InvokeMethod("nSetModule", "A/R");

                    // Set task ID to oSS
                    int TaskID = (int)oSS.InvokeMethod("nLookupTask", "SO_SalesOrder_ui");
                    retVal = oSS.InvokeMethod("nSetProgram", TaskID);

                    // SafeProcess just checks the last error message value and prints it to
                    // the debug output. In the future, I will have it post this to an error log
                    SafeProcess(oSS.InvokeMethod("nSetProgram", TaskID), oSS);

                    // Loop through the orders
                    // This saves time over opening the oSS and PVX object over and over again
                    foreach (SalesOrder salesorder in this.SalesOrders())
                    {
                        Debug.WriteLine(">>Changing " + salesorder.SalesOrderNo + "<<");
                        // Open object for sales order
                        using (DispatchObject so_order = new DispatchObject(pvx.InvokeMethod("NewObject", "SO_SalesOrder_bus", oSS.GetObject())))
                        {

                            SafeProcess(so_order.InvokeMethod("nSetKey", salesorder.SalesOrderNo), so_order);

                            SafeProcess(so_order.InvokeMethod("nSetValue", "ShipVia$", shipvia), so_order);

                            retVal = so_order.InvokeMethod("nWrite");
                            if ((int)retVal != 1)
                            {
                                Debug.WriteLine(so_order.GetProperty("sLastErrorMsg"));
                            }
                        }

                        Debug.WriteLine("<<End Changing Sales Order<<");
                    }
                }

            }
        }
        public List<Invoice> Invoice()
        {
            List<Invoice> invoices = new List<Invoice>();

            Microsoft.Win32.RegistryKey SageKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\SOTAMAS90");
            string InitPath = SageKey.GetValue("Directory") + @"\Home";
            object retVal;

            // Open Initial PVX object
            using (DispatchObject pvx = new DispatchObject("ProvideX.Script"))
            {
                // Set MAS90 Path
                pvx.InvokeMethod("Init", InitPath);

                // Set up my Session connection
                using (DispatchObject oSS = new DispatchObject(pvx.InvokeMethod("NewObject", "SY_Session")))
                {
                    // Set Username and password
                    // Set Company, and set the date
                    // Basic setup for the oSS object
                    oSS.InvokeMethod("nSetUser", Username, Password);
                    oSS.InvokeMethod("nLogon");
                    oSS.InvokeMethod("nSetCompany", Company);
                    oSS.InvokeMethod("nSetDate", "S/O", DateTime.Today.ToString("yyyyMMdd"));
                    oSS.InvokeMethod("nSetModule", "A/R");

                    // Set task ID to oSS
                    int TaskID = (int)oSS.InvokeMethod("nLookupTask", "SO_Invoice_ui");
                    retVal = oSS.InvokeMethod("nSetProgram", TaskID);

                    SafeProcess(oSS.InvokeMethod("nSetProgram", TaskID), oSS);

                    using (DispatchObject so_invoice = new DispatchObject(pvx.InvokeMethod("NewObject", "SO_Invoice_bus", oSS.GetObject())))
                    {
                        object[] batchnum = new object[] { BatchNo, "N", BatchDescription };

                        if (BatchNo == "")
                        {
                            SafeProcess(so_invoice.InvokeMethodByRef("nSelectNewBatch", batchnum), so_invoice);
                            BatchNo = (string)batchnum[0];

                        }
                        else
                        {
                            SafeProcess(so_invoice.InvokeMethod("nSelectBatch", BatchNo), so_invoice);
                        }

                        foreach (SalesOrder order in SalesOrders())
                        {
                            string InvoiceNumber = string.Empty;
                            object[] nextInvoiceNum = new object[] { "OrderNo$" };
                            SafeProcess(so_invoice.InvokeMethodByRef("nGetNextInvoiceNo", nextInvoiceNum), so_invoice);

                            InvoiceNumber = (string)nextInvoiceNum[0];

                            SafeProcess(so_invoice.InvokeMethod("nSetKey", InvoiceNumber), so_invoice);

                            SafeProcess(so_invoice.InvokeMethod("nsetvalue", "SalesOrderNo$", order.SalesOrderNo), so_invoice);

                            //SafeProcess(so_invoice.InvokeMethod("nsetvalue", "ShipWeight$", "13.37"), so_invoice);

                            using (DispatchObject so_line = new DispatchObject(so_invoice.GetProperty("oLines")))
                            {
                                SafeProcess(so_line.InvokeMethod("nCopyLinesFromSalesOrder", order.SalesOrderNo, "Y"), so_line);
                                //retVal = so_line.InvokeMethod("nCopyLinesFromSalesOrder", SalesOrderNo, "Y");

                                SafeProcess(so_line.InvokeMethod("nMoveFirst"), so_line);

                                string ItemType = string.Empty;

                                object[] getValueParam3 = new object[] { "ItemType$", "" };

                                SafeProcess(so_line.InvokeMethodByRef("nGetValue", getValueParam3), so_line);


                                int EOF = (int)so_line.GetProperty("nEOF");
                                int ShippedQty = 0;

                                do
                                {
                                    int orderquantity = 0;
                                    object[] orderedQty = new object[] { "QuantityOrdered", orderquantity };

                                    SafeProcess(so_line.InvokeMethodByRef("nGetValue", orderedQty), so_line);
                                    ShippedQty = (int)orderedQty[1];

                                    //object[] getValueParam2 = new object[] { "ItemCodeDesc$", "" };
                                    // Read Into the object
                                    //SafeProcess(so_line.InvokeMethodByRef("nGetValue", getValueParam2), so_line);

                                    SafeProcess(so_line.InvokeMethod("nsetvalue", "QuantityShipped", ShippedQty), so_line);

                                    SafeProcess(so_line.InvokeMethod("nWrite"), so_line);

                                    SafeProcess(so_line.InvokeMethod("nMoveNext"), so_line);

                                    EOF = (int)so_line.GetProperty("nEOF");

                                } while (EOF != 1);

                                SafeProcess(so_invoice.InvokeMethod("nsetvalue", "FreightAmt", order.Freight), so_invoice);
                                SafeProcess(so_invoice.InvokeMethod("nsetvalue", "NumberofPackages", order.TrackingNumbers.Count()), so_invoice);

                                SafeProcess(so_invoice.InvokeMethod("nWrite"), so_invoice);

                                invoices.Add(new Invoice(InvoiceNumber));
                                Debug.WriteLine("Invoice Created: " + InvoiceNumber);

                            }
                        }

                    }


                }
            }

            return invoices;
        }

        private static int SafeProcess(object Method, DispatchObject Object)
        {
            object retVal;

            retVal = Method;
            if ((int)retVal != 1)
            {
                // Output error to error log in future
                //Debug.WriteLine(Object.GetProperty("sLastErrorMsg").ToString());
            }

            return (int)retVal;
        }

    }

}