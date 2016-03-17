using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace SageIntegration
{
    class Core
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

        public class SageObject
        {
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
                    Debug.Write(ErrorMessage);
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
