using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace SageIntegration
{
    public static class TextUtilities
    {
        public static string Left(this String input, int length)
        {
            if (input.Length > length)
            {
                Debug.WriteLine("Shortening <" + input + "> to " + length + " characters");
            }
            return (input.Length < length) ? input : input.Substring(0, length);
        }
    }
}
