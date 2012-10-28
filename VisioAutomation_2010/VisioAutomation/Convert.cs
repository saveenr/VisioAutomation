using VA=VisioAutomation;

namespace VisioAutomation
{
    public static class Convert
    {
        private const string quote = "\"";
        private const string quotequote = "\"\"";

        public static short BoolToShort(bool b)
        {
            return b ? ((short)1) : ((short)0);
        }

        public static string BoolToFormula(bool b)
        {
            return b ? "1" : "0";
        }

        public static bool DoubleToBool(double d)
        {
            return d != 0;
        }

        /// <summary>
        /// Converts a short value to bool
        /// </summary>
        /// <param name="v"></param>
        /// <returns></returns>
        public static bool ShortToBool(short v)
        {
            // if v is 0 then false
            // if v != 0 then true
            return v != 0;
        }
        
        /// <summary>
        /// Properly quotes a string being used as a formula
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string StringToFormulaString(string s)
        {
            if (s == null)
            {
                throw new System.ArgumentNullException("s");
            }

            string result = System.String.Format("\"{0}\"", s.Replace(quote, quotequote));
            return result;
        }

        public static string FormulaStringToString(string formula)
        {
            if (formula == null)
            {
                throw new System.ArgumentNullException("formula");
            }

            // Initialize the converted formula from the value passed in.
            string output_string = formula;

            // Check if this formula value is a quoted string.
            // If it is, remove extra quote characters.
            if (output_string.StartsWith(quote) &&
                output_string.EndsWith(quote))
            {

                // Remove the wrapping quote characters as well as any
                // extra quote marks in the body of the string.
                output_string = output_string.Substring(1, (output_string.Length - 2));
                output_string = output_string.Replace(quotequote, quote);
            }

            return output_string;
        }
    }
}