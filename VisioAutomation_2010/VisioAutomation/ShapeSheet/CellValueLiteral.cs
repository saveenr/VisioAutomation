namespace VisioAutomation.ShapeSheet
{
    /// <summary>
    /// CellValueLiteral is used in those cases where you want a caller to provide an int, double, bool, or string value to be used as a formula
    /// In the case of string inputs, no special escaping of strings is performed. The caller must do any escaping.
    /// This struct allows you to have one method that takes multiple types as a parameter (via implicit conversion) and is
    /// similar to the use of XName in System.Linq.Xml.
    /// 
    /// IMPORTANT: The value stored is always a string. Any input will be converted to a string.
    /// </summary>
    public struct CellValueLiteral
    {
        private readonly string _stringval;
        public string Value => this._stringval;
        public bool HasValue => this._stringval != null;
        public override string ToString() => this.Value;

        public CellValueLiteral(string value)
        {
            this._stringval = value;
        }

        public CellValueLiteral(int value)
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            this._stringval = value.ToString(culture);
        }

        public CellValueLiteral(double value)
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            this._stringval = value.ToString(culture);
        }

        public CellValueLiteral(bool value)
        {
            this._stringval = value ? "TRUE" : "FALSE";
        }

        public static implicit operator CellValueLiteral(string value)
        {
            return new CellValueLiteral(value);
        }

        public static implicit operator CellValueLiteral(int value)
        {
            return new CellValueLiteral(value);
        }

        public static implicit operator CellValueLiteral(double value)
        {
            return new CellValueLiteral(value);
        }

        public static implicit operator CellValueLiteral(bool value)
        {
            return new CellValueLiteral(value);
        }

        public static string EncodeValue(string text)
        {
            return EncodeValue(text, true);
        }

        public static string EncodeValue(string text, bool autoquote)
        {
            // Some cells are very pick about values being quoted
            // This method is a reasonable way of getting values quoted smartly
            // and avoids quoting multiple times

            // Rules are simple:
            // - passthrough null values
            // - passthrough empty values
            // - passthrough values that begin with = - it is assumed the such strings have been carefully crafter
            // - passthrough values that begin with " - it is assumed the such strings are already correct
            // - finally (of noqutoe==false) replace " with "" and surround with " the result

            if (text == null)
            {
                return text;
            }

            if (text.Length == 0)
            {
                return text;
            }

            if (text[0] == '\"')
            {
                return text;
            }

            if (text[0] == '=')
            {
                return text;
            }
            
            // if the caller wants to force the content to a formula string
            // then do so: escape internal double quotes and then wrap in double quotes
            if (autoquote)
            {
                string str_quoted = text.Replace("\"", "\"\"");
                str_quoted = string.Format("\"{0}\"", str_quoted);
                return str_quoted;
            }

            // For all other cases, just return the input string
            return text;
        }


    }
}