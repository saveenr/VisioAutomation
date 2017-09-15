namespace VisioAutomation.Utilities
{
    public static class Convert
    {
        private const string doublequote = "\"";
        private const string doublequote_x2 = "\"\"";

        public static string FormulaEncodeSmart(string s)
        {
            if (s == null)
            {
                throw new System.ArgumentNullException(nameof(s));
            }

            // if its empty or begins with '=' return it as is
            if (s.Length == 0 || s[0]=='=')
            {
                return s;
            }

            var result_quote_escaped = s.Replace(Convert.doublequote, Convert.doublequote_x2);
            string result = string.Format("\"{0}\"", result_quote_escaped);

            return result;
        }
    }
}