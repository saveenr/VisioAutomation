namespace VisioAutomation.ShapeSheet
{
    public struct CellData
    {
        public string Value { get; }

        public CellData(string formula)
            : this()
        {
            this.Value = formula;
        }

        public override string ToString()
        {
            var formula_string = this.Value ?? "null";
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            var format = "\"{0}\"";
            return string.Format(culture,format, formula_string);
        }

        public static implicit operator CellData(CellValueLiteral formula)
        {
            return new CellData(formula.Value);
        }

        public static implicit operator CellData(string formula)
        {
            return new CellData(formula);
        }

        public static implicit operator CellData(int formula)
        {
            var cv = (CellValueLiteral) formula;
            return new CellData(cv.Value);
        }

        public static implicit operator CellData(double formula)
        {
            var cv = (CellValueLiteral)formula;
            return new CellData(cv.Value);
        }

        public static implicit operator CellData(bool formula)
        {
            var cv = (CellValueLiteral)formula;
            return new CellData(cv.Value);
        }
    }
}