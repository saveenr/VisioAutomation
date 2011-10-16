using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet.Update
{
    public class UpdateBase<T> where T : struct
    {
        protected FormulaData<T> FormulaData { get; private set; }
        protected ResultData<T> ResultData { get; private set; }

        protected UpdateBase()
        {
            this.FormulaData = new FormulaData<T>();
            this.ResultData = new ResultData<T>();
        }

        protected UpdateBase(int fcapacity, int rcapacity)
        {
            this.FormulaData = new FormulaData<T>(fcapacity);
            this.ResultData = new ResultData<T>(rcapacity);
        }

        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        public IVisio.VisGetSetArgs ResultFlags
        {
            get
            {
                return get_common_flags();
            }
        }

        public IVisio.VisGetSetArgs FormulaFlags
        {
            get
            {
                var common_flags = get_common_flags();
                var formula_flags = (short) IVisio.VisGetSetArgs.visSetUniversalSyntax;
                var combined_flags = (short)common_flags | formula_flags;
                return (IVisio.VisGetSetArgs) combined_flags;
            }
        }

        private IVisio.VisGetSetArgs get_common_flags()
        {
            IVisio.VisGetSetArgs f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            IVisio.VisGetSetArgs f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = (short)f_bg | (short)f_tc;
            return (IVisio.VisGetSetArgs)flags;
        }

        public void SetFormula(T streamitem, FormulaLiteral literal)
        {
            this.FormulaData.Set(streamitem,literal);
        }

        public void SetFormulaIgnoreNull(T streamitem, ShapeSheet.FormulaLiteral f)
        {
            if (f.HasValue)
            {
                this.SetFormula(streamitem, f);
            }
        }

        public void SetResult(T streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this.ResultData.Set(streamitem,value,unitcode);
        }
    }
}