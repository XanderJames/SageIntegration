using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace SageIntegration
{
    public class Core
    {
        public static bool CanConnectToSage()
        {
            bool result = false;
            Microsoft.Win32.RegistryKey SageKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\SOTAMAS90");
            string InitPath = SageKey.GetValue("Directory") + @"\Home";

            DispatchObject Test_PVX = null;
            DispatchObject Test_OSS = null;

            try
            {
                // If these objects fail, we need to manually clean them up
                // wrapping them in a using does not clean dispose them
                Test_PVX = new DispatchObject("ProvideX.Script");
                Test_PVX.InvokeMethod("Init", InitPath);
                Test_OSS = new DispatchObject(Test_PVX.InvokeMethod("NewObject", "SY_Session"));
                // If we get this far, the connection to Sage was a success
                result = true;

            }
            catch (System.Reflection.TargetInvocationException)
            {
                // We can't connect to Sage 100. Let the user know this.
                // We are going to return
            }
            finally
            {
                /*
                 * We have to force these objects to dispose
                 */

                if (Test_PVX != null)
                {
                    Test_PVX.Dispose();
                }

                if (Test_OSS != null)
                {
                    Test_OSS.Dispose();
                }

            }
            return result;
        }
        public class Credentials
        {
            public string Username { get; set; }
            public string Password { get; set; }

            public Credentials(string username, string password)
            {
                Username = username;
                Password = password;
            }
        }

        public class Company
        {
            public string Code { get; set; }
            public string Description { get; set; }
            public Company(string code, string description = "")
            {
                Code = code;
                Description = description;
            }
        }

        public static List<Company> CompanyList()
        {
            List<Company> Companies = new List<Company>();
            Microsoft.Win32.RegistryKey SageKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\SOTAMAS90");
            string InitPath = SageKey.GetValue("Directory") + @"\Home";
            object retVal = new object();

            // See if we can make a connection to Sage
            if (Core.CanConnectToSage() == false)
            {
                // We can't connect to Sage 100
                return null;
            }

            // Open Initial PVX object
            using (DispatchObject pvx = new DispatchObject("ProvideX.Script"))
            {
                pvx.InvokeMethod("Init", InitPath);

                // Set up my Session connection
                using (DispatchObject oSS = new DispatchObject(pvx.InvokeMethod("NewObject", "SY_Session")))
                {
                    // Split on chr 138 - Š
                    string[] CompanyList = oSS.GetProperty("sCompanyList").ToString().Split(new char[] { 'Š' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string company in CompanyList)
                    {
                        Companies.Add(
                            new Company(
                                company.Split(new string[] { "  " }, StringSplitOptions.None)[0], 
                                company.Split(new string[] { "  " }, StringSplitOptions.None)[1].Replace(")", "").Replace("(", "")));
                    }
                }
            }

            return Companies;
        }

        public class SageObject
        {
            // All errors are passed into this event
            // Handle these error messages however you want
            // in your application => Logging or a message box
            public static event RaiseEvent SageError;
            public delegate void RaiseEvent(string EventText);

            public int ReturnCode { get; set; }
            public string LastErrorMessage { get; set; }

            public static SageObject Process(object Method, DispatchObject Object)
            {
                object retVal;
                string ErrorMessage = string.Empty;
                retVal = Method;

                if ((int)retVal != 1)
                {
                    ErrorMessage = Object.GetProperty("sLastErrorMsg").ToString();
                    if (SageError != null) { SageError(ErrorMessage); }
                }

                return new SageObject(ErrorMessage, (int)retVal);
            }

            public SageObject(string LastErrorMessage, int ReturnCode)
            {
                this.LastErrorMessage = LastErrorMessage;
                this.ReturnCode = ReturnCode;
            }
        }

        public enum ReturnCode
        {
            SuccessWithError = -1,
            Failed = 0,
            Success = 1,
            SuccessWithMinorError = 2,
            ModuleSet = 100007
        }
    }
}
