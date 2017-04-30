using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public struct CellData
    {
        public string Formula { get; }
        public string Result { get; }

        public CellData(string formula, string result)
            : this()
        {
            this.Formula = formula;
            this.Result = result;
        }

        public override string ToString()
        {
            var formula_string = this.Formula ?? "null";
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            var format = "(\"{0}\",{1})";
            return string.Format(invariant_culture,format, formula_string, this.Result);
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