using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public struct CellData<TResult>
    {
        public VA.ShapeSheet.FormulaLiteral Formula { get; set; }
        public TResult Result { get; set; }

        public CellData(VA.ShapeSheet.FormulaLiteral formula, TResult result)
            : this()
        {
            this.Formula = formula;
            this.Result = result;
        }

        /// <summary>
        /// Returns whether there is a formula stored. If the formula is set to null then it counts as not having a formula
        /// </summary>
        public bool HasFormula
        {
            get { return this.Formula.HasValue; }
        }

        public override string ToString()
        {
            var fs = (this.HasFormula) ? string.Format("\"{0}\"", this.Formula) : "null";
            var rs = this.GetResultAsString();
            return string.Format("({0},{1})", fs, rs);
        }

        private string GetResultAsString()
        {
            return this.Result.ToString();
        }

        internal void SetResult(TResult result)
        {
            this.Result = result;
        }

        internal void SetFormula(VA.ShapeSheet.FormulaLiteral formula)
        {
            this.Formula = formula;
        }

        public static implicit operator CellData<TResult>(VA.ShapeSheet.FormulaLiteral formula)
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

    }
}