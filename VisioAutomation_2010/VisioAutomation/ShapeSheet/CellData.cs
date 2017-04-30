namespace VisioAutomation.ShapeSheet
{
    public struct CellData
    {
        public CellValueLiteral Formula { get; }
        public string Result { get; }

        public CellData(CellValueLiteral formula, string result)
            : this()
        {
            this.Formula = formula;
            this.Result = result;
        }

        public override string ToString()
        {
            var formula_string = (this.Formula.HasValue) ? this.Formula.Value : "null";
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            var format = "(\"{0}\",{1})";
            return string.Format(invariant_culture,format, formula_string, this.Result);
        }

        public static implicit operator CellData(CellValueLiteral formula)
        {
            return new CellData(formula,default(string));
        }

        public static implicit operator CellData(string formula)
        {
            return new CellData(formula, default(string));
        }

        public static implicit operator CellData(int formula)
        {
            return new CellData(formula, default(string));
        }

        public static implicit operator CellData(double formula)
        {
            return new CellData(formula, default(string));
        }

        public static implicit operator CellData(bool formula)
        {
            return new CellData(formula, default(string));
        }
    }
}