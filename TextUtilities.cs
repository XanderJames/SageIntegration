using System.Text.RegularExpressions;

namespace SageIntegration
{
    public static class TextUtilities
    {
        public static string Left(this string input, int length)
        {
            return (input.Length < length) ? input : input.Substring(0, length);
        }

        public static string LimitString(string Input, int Length)
        {
            if (Input == string.Empty)
            {
                return Input;
            }

            Input = Regex.Replace(Input, @"[^\xA0-\xFF\u0000-\u007E€]", "?").Trim();

            if (Input.Length < Length)
            {
                return Input;
            }
            else
            {
                return Input.Substring(0, Length);
            }

        }
    }
}
