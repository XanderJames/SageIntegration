using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SageIntegration
{
    public class ItemCode
    {
        public string Code { get; set; }
        public string Description { get; set; }

        public string ProductLine { get; set; }
        public decimal Weight { get; set; }
        public decimal Volume { get; set; }
        public string UOM_Standard { get; set; }
        public string Vendor { get; set; }
        public string DefaultWarehouse { get; set; }
        
        public static void Copy(
            Core.Company Company, Core.Credentials Credentials, 
            string SourceCode, string DestinationCode, 
            bool CopyItemVendor = true, bool CopyAliasItems = true, bool UpdateBillofMaterials = false)
        {
            // Grab init path from the registry
            Microsoft.Win32.RegistryKey SageKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\SOTAMAS90");
            string InitPath = SageKey.GetValue("Directory") + @"\Home";
            object retVal = new object();

            // See if we can make a connection to Sage
            if (Core.CanConnectToSage() == false)
            {
                // We can't connect to Sage 100
            }

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


                    if (Core.SageObject.Process(oSS.InvokeMethod("nSetCompany", Company.Code), oSS).ReturnCode == (int)Core.ReturnCode.Failed)
                    {
                        // I'm returning a null here to break anything
                        // that doesn't expect a null. This isn't something
                        // I should run into unless I'm using incorrect
                        // company codes.
                        return;
                    }

                    if (Core.SageObject.Process(oSS.InvokeMethod("nSetUser", Credentials.Username, Credentials.Password), oSS).ReturnCode == (int)Core.ReturnCode.Failed)
                    {
                        // Login Failure
                        // Check Username and Password
                        return;
                    }
                    oSS.InvokeMethod("nLogon");


                    oSS.InvokeMethod("nSetDate", "C/I", DateTime.Today.ToString("yyyyMMdd"));
                    oSS.InvokeMethod("nSetModule", "C/I");

                    // Set task ID to oSS
                    int TaskID = (int)oSS.InvokeMethod("nLookupTask", "CI_ItemCode_ui");
                    Core.SageObject.Process(oSS.InvokeMethod("nSetProgram", TaskID), oSS);

                    // Connect to ItemCode bus, create a new item, then copy from it
                    using (DispatchObject ci_item = new DispatchObject(pvx.InvokeMethod("NewObject", "CI_ItemCode_bus", oSS.GetObject())))
                    {
                        Core.SageObject.Process(ci_item.InvokeMethod("nSetKey", new object[] { DestinationCode }), ci_item);
                        Core.SageObject.Process(ci_item.InvokeMethod("nCopyFrom", new object[] { SourceCode }, true, true, true), ci_item);

                        if (Core.SageObject.Process(ci_item.InvokeMethod("nWrite"), ci_item).ReturnCode == (int)Core.ReturnCode.Failed)
                        {
                            return;
                        }
                    }

                    // If true, copy Bill of Materials
                    if (UpdateBillofMaterials)
                    {
                        oSS.InvokeMethod("nSetDate", "B/M", DateTime.Today.ToString("yyyyMMdd"));
                        oSS.InvokeMethod("nSetModule", "B/M");

                        TaskID = (int)oSS.InvokeMethod("nLookupTask", "BM_Bill_ui");
                        Core.SageObject.Process(oSS.InvokeMethod("nSetProgram", TaskID), oSS);

                        using (DispatchObject bm_bill = new DispatchObject(pvx.InvokeMethod("NewObject", "BM_Bill_bus", oSS.GetObject())))
                        {
                            Core.SageObject.Process(bm_bill.InvokeMethod("nSetValue", "BILLNO$", DestinationCode), bm_bill);
                            Core.SageObject.Process(bm_bill.InvokeMethod("nSetKey"), bm_bill);

                            Core.SageObject.Process(bm_bill.InvokeMethod("nSetCopyBillNo", new object[] { SourceCode, "" }), bm_bill);
                            Core.SageObject.Process(bm_bill.InvokeMethod("nCopyFrom"), bm_bill);
                            Core.SageObject.Process(bm_bill.InvokeMethod("nWrite"), bm_bill);
                        }

                    }

                }

            }

        }

        public static List<string> List(Core.Company Company, Core.Credentials Credentials)
        {

            // NOT WORKING
            // NO IDEA WHY - RETURNS NOTHING :<
            List<string> result = new List<string>();

            // Grab init path from the registry
            Microsoft.Win32.RegistryKey SageKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\SOTAMAS90");
            string InitPath = SageKey.GetValue("Directory") + @"\Home";
            object retVal = new object();

            // See if we can make a connection to Sage
            if (Core.CanConnectToSage() == false)
            {
                // We can't connect to Sage 100
            }

            // Open Initial PVX object
            using (DispatchObject pvx = new DispatchObject("ProvideX.Script"))
            {
                pvx.InvokeMethod("Init", InitPath);

                // Set up my Session connection
                using (DispatchObject oSS = new DispatchObject(pvx.InvokeMethod("NewObject", "SY_Session")))
                {
                    if (Core.SageObject.Process(oSS.InvokeMethod("nSetCompany", Company.Code), oSS).ReturnCode == (int)Core.ReturnCode.Failed)
                    {
                        // I'm returning a null here to break anything
                        // that doesn't expect a null. This isn't something
                        // I should run into unless I'm using incorrect
                        // company codes.
                        return result;
                    }

                    if (Core.SageObject.Process(oSS.InvokeMethod("nSetUser", Credentials.Username, Credentials.Password), oSS).ReturnCode == (int)Core.ReturnCode.Failed)
                    {
                        // Login Failure
                        // Check Username and Password
                        return result;
                    }
                    oSS.InvokeMethod("nLogon");



                    /* PULLED WORKING CODE FROM OTHER PROJECT
                    oSS.InvokeMethod("nSetDate", "A/R", DateTime.Today.ToString("yyyyMMdd"));
                    oSS.InvokeMethod("nSetModule", "A/R");
                    
                    int TaskID = (int)oSS.InvokeMethod("nLookupTask", "AR_Customer_ui");
                    Core.SageObject.Process(oSS.InvokeMethod("nSetProgram", TaskID), oSS);

                    using (DispatchObject ar_customer = new DispatchObject(pvx.InvokeMethod("NewObject", "AR_Customer_svc", oSS.GetObject())))
                    {
                        String columns = "CustomerName$";
                        String keys = "CustomerNo$";
                        String returnFields = "CustomerNo$";
                        String returnAccountKeys = "CustomerNo$";
                        String whereClause = "";

                        Object[] getResultSetParams = new Object[] { columns, keys, returnFields, returnAccountKeys, whereClause, "", "" };

                        object retTest = ar_customer.InvokeMethodByRef("nGetResultSets", getResultSetParams);

                    }
                     
                     */

                    oSS.InvokeMethod("nSetDate", "C/I", DateTime.Today.ToString("yyyyMMdd"));
                    oSS.InvokeMethod("nSetModule", "C/I");

                    using (DispatchObject ci_item = new DispatchObject(pvx.InvokeMethod("NewObject", "CI_ItemCode_svc", oSS.GetObject())))
                    {
                        String columns = "ItemCode$";
                        String keys = "ItemCode$";
                        String returnFields = "";
                        String returnAccountKeys = "";
                        String whereClause = "";

                        Object[] getResultSetParams = new Object[] { columns, keys, returnFields, returnAccountKeys, whereClause, "", "" };

                        //"ItemCode$", "ItemCode$", TestString, "ItemType$=" + (char)34 + "1" + (char)34 + " AND UseInSO$=" + (char)34 + "Y" + (char)34, "", "")
                        Core.SageObject.Process(ci_item.InvokeMethod("nGetResultSets", getResultSetParams), ci_item);
                    }

                    return result;
                }
            }
        }

    }
}
