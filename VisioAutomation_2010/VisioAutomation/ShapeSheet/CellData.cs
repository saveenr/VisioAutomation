namespace VisioAutomation.ShapeSheet
{
    public struct CellData
    {
        public string ValueF { get; }
        public string ValueR { get; }

        public CellData(string formula, string result)
            : this()
        {
            this.ValueF = formula;
            this.ValueR = result;
        }

        public override string ToString()
        {
            var formula_string = this.ValueF ?? "null";
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            var format = "(\"{0}\",{1})";
            return string.Format(culture,format, formula_string, this.ValueR);
        }

        public static implicit operator CellData(CellValueLiteral formula)
        {
            return new CellData(formula.Value, default(string));
        }

        public static implicit operator CellData(string formula)
        {
            return new CellData(formula, default(string));
        }

        public static implicit operator CellData(int formula)
        {
            var cv = (CellValueLiteral) formula;
            return new CellData(cv.Value, default(string));
        }

        public static implicit operator CellData(double formula)
        {
            var cv = (CellValueLiteral)formula;
            return new CellData(cv.Value, default(string));
        }

        public static implicit operator CellData(bool formula)
        {
            var cv = (CellValueLiteral)formula;
            return new CellData(cv.Value, default(string));
        }
    }
}