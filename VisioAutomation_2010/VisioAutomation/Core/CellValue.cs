namespace VisioAutomation.Core
{
    /// <summary>
    /// CellValueLiteral is used in those cases where you want a caller to provide an int, double, bool, or string value to be used as a formula
    /// In the case of string inputs, no special escaping of strings is performed. The caller must do any escaping.
    /// This struct allows you to have one method that takes multiple types as a parameter (via implicit conversion) and is
    /// similar to the use of XName in System.Linq.Xml.
    /// 
    /// IMPORTANT: The value stored is always a string. Any input will be converted to a string.
    /// </summary>
    public struct CellValue
    {
        private readonly string _stringval;
        private static string str_const_true = "TRUE";
        private static string str_const_false = "FALSE";
        private static char _char_equalssign = '=';
        private static char _char_doublequote = '\"';
        private static string _str_doublequote = "\"";
        private static string _str_twodoublequotes = "\"\"";

        public string Value => this._stringval;
        public bool HasValue => this._stringval != null;
        public override string ToString() => this.Value;

        public CellValue(string value)
        {
            this._stringval = value;
        }

        public CellValue(int value)
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            this._stringval = value.ToString(culture);
        }

        public CellValue(double value)
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            this._stringval = value.ToString(culture);
        }

        public CellValue(bool value)
        {
            this._stringval = value ? str_const_true : str_const_false;
        }

        public static implicit operator CellValue(string value)
        {
            return new CellValue(value);
        }

        public static implicit operator CellValue(int value)
        {
            return new CellValue(value);
        }

        public static implicit operator CellValue(double value)
        {
            return new CellValue(value);
        }

        public static implicit operator CellValue(bool value)
        {
            return new CellValue(value);
        }

        public static string EncodeValue(string text)
        {
            return EncodeValue(text, true);
        }

        public static string EncodeValue(string text, bool quote)
        {
            // Some cells are very pick about values being quoted
            // This method is a reasonable way of getting values quoted smartly
            // and avoids quoting multiple times

            // Rules are executed in this order:
            // - null or empty or begins with = -> don't change
            // - begins with " and ends with " -> don't change (assume caller has carefully crafted it)
            // - if (quote flag on) - quote the string
            //      turn all " to ""
            //      add beginning " and ending "
            // - don't change
            
            //  if noquote==false) replace " with "" and surround with " the result

            if (string.IsNullOrEmpty(text) || text[0] == _char_equalssign)
            {
                return text;
            }

            // it's quoted already, just return it
            if (text[0] == _char_doublequote && text[text.Length-1]==_char_doublequote)
            {
                return text;
            }

            if (quote)
            {

                string str_quoted = text.Replace(_str_doublequote, _str_twodoublequotes);
                str_quoted = string.Format("\"{0}\"", str_quoted);
                return str_quoted;
            }

            // For all other cases, just return the input string
            return text;
        }

        internal bool ValidateValue(bool quote_required)
        {
            string text = this.Value;

            if (text == null)
            {
                return true;
            }

            if (text.Length == 0)
            {
                return true;
            }

            if (text[0] == _char_equalssign)
            {
                return true;
            }

            if (text[0] == _char_doublequote)
            {
                if (text[text.Length - 1] != _char_doublequote)
                {
                    return false;
                }
            }
            else
            {
                if (quote_required)
                {
                    return false;
                }
            }
            return true;
        }
    }
}