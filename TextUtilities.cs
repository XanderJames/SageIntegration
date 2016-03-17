namespace SageIntegration
{
    public static class TextUtilities
    {
        public static string Left(this string input, int length)
        {
            return (input.Length < length) ? input : input.Substring(0, length);
        }
    }
}
