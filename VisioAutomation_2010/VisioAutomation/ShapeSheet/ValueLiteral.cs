using VisioAutomation.Utilities;

namespace VisioAutomation.ShapeSheet
{
    /// <summary>
    /// FormulaLiteral is used in those cases where you want a caller to provide an int, double, bool, or string value to be used as a formula
    /// In the case of string inputs, no special escaping of strings is performed. The caller must do any escaping.
    /// This struct allows you to have one method that takes multiple types as a parameter (via implicit conversion) and is
    /// similar to the use of XName in SXL.
    /// 
    /// The value stored is always a string. Any input will be converted to a string.
    /// </summary>
    public struct ValueLiteral
    {
        private readonly string _stringval;

        private ValueLiteral(string s)
        {
            this._stringval = s;
        }

        public string Value => this._stringval;

        public bool HasValue => this._stringval != null;

        public override string ToString() => this.Value;

        public static implicit operator ValueLiteral(string value)
        {
            return new ValueLiteral(value);
        }

        public static implicit operator ValueLiteral(int value)
        {
            var formula = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return new ValueLiteral(formula);
        }

        public static implicit operator ValueLiteral(double value)
        {
            var formula = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return new ValueLiteral(formula);
        }

        public static implicit operator ValueLiteral(bool value)
        {
            var formula = value ? "1" : "0";
            return new ValueLiteral(formula);
        }

        public string Encode()
        {
            if (!this.HasValue)
            {
                throw new System.ArgumentException("No Value to Encode");
            }
            return Convert.StringToFormulaString(this.Value);
        }
    }
}