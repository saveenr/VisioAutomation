using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public struct CellData<TResult>
    {
        public VA.ShapeSheet.FormulaLiteral Formula { get; private set; }
        public TResult Result { get; private set; }

        public CellData(VA.ShapeSheet.FormulaLiteral formula, TResult result)
            : this()
        {
            this.Formula = formula;
            this.Result = result;
        }

        public override string ToString()
        {
            var fs = (this.Formula.HasValue) ? string.Format("\"{0}\"", this.Formula.Value) : "null";
            var rs = this.Result.ToString();
            return string.Format("({0},{1})", fs, rs);
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
    }
}