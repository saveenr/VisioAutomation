namespace VisioAutomation.ShapeSheet
{
    public struct CellData<TResult>
    {
        public FormulaLiteral Formula { get; }
        public TResult Result { get; }

        public CellData(FormulaLiteral formula, TResult result)
            : this()
        {
            this.Formula = formula;
            this.Result = result;
        }

        public override string ToString()
        {
            var fs = (this.Formula.HasValue) ? this.Formula.Value : "null";
            return string.Format(System.Globalization.CultureInfo.InvariantCulture,"(\"{0}\",{1})", fs, this.Result);
        }

        public static implicit operator CellData<TResult>(FormulaLiteral formula)
        {
            return new CellData<TResult>(formula,default(TResult));
        }

        public static implicit operator CellData<TResult>(string formula)
        {
            return new CellData<TResult>( formula, default(TResult));
        }

        public static implicit operator CellData<TResult>(int formula)
        {
            return new CellData<TResult>(formula, default(TResult));
        }

        public static implicit operator CellData<TResult>(double formula)
        {
            return new CellData<TResult>(formula, default(TResult));
        }

        public static implicit operator CellData<TResult>(bool formula)
        {
            return new CellData<TResult>(formula, default(TResult));
        }

        public static CellData<TResult>[] Combine<TResult>(TResult[] results, string[] formulas)
        {
            var combined_data = new ShapeSheet.CellData<TResult>[results.Length];
            for (int i = 0; i < results.Length; i++)
            {
                combined_data[i] = new ShapeSheet.CellData<TResult>(formulas[i], results[i]);
            }
            return combined_data;
        }
    }
}