using VA = VisioAutomation;

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
    public struct FormulaLiteral
    {
        private readonly string formula_string;

        private FormulaLiteral(string s)
        {
            this.formula_string = s;
        }

        public string Value => this.formula_string;

        public bool HasValue => this.formula_string != null;

        public static implicit operator FormulaLiteral(string value)
        {
            return new FormulaLiteral(value);
        }

        public static implicit operator FormulaLiteral(int value)
        {
            var formula = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return new FormulaLiteral(formula);
        }

        public static implicit operator FormulaLiteral(double value)
        {
            var formula = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return new FormulaLiteral(formula);
        }

        public static implicit operator FormulaLiteral(bool value)
        {
            var formula = value ? "1" : "0";
            return new FormulaLiteral(formula);
        }

        public override string ToString()
        {
            return this.Value;
        }

        public string Encode()
        {
            if (!this.HasValue)
            {
                throw new AutomationException("No Value to Encode");
            }
            return Convert.StringToFormulaString(this.Value);
        }
    }
}