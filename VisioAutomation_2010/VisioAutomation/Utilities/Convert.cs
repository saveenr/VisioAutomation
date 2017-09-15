namespace VisioAutomation.Utilities
{
    public static class Convert
    {
        private const string doublequote = "\"";
        private const string doublequote_x2 = "\"\"";

        public static string FormulaEncodeSmart(string text)
        {
            if (text == null)
            {
                throw new System.ArgumentNullException(nameof(text));
            }

            if (text.Length == 0)
            {
                return text;
            }

            if (text[0] == '\"')
            {
                return text;
            }

            var result_quote_escaped = text.Replace(Convert.doublequote, Convert.doublequote_x2);
            string result = string.Format("\"{0}\"", result_quote_escaped);

            return result;
        }
    }
}