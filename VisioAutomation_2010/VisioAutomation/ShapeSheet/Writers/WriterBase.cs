using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public abstract class WriterBase<TStreamType, TValue>
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        public List<TStreamType> StreamItems;
        public List<TValue> ValueItems;

        public void Clear()
        {
            this.StreamItems.Clear();
            this.ValueItems.Clear();
        }

        protected WriterBase()
        {
            this.StreamItems = new List<TStreamType>();
            this.ValueItems = new List<TValue>();
        }

        protected WriterBase(int capacity)
        {
            this.StreamItems = new List<TStreamType>(capacity);
            this.ValueItems = new List<TValue>(capacity);
        }

        protected IVisio.VisGetSetArgs ResultFlags
        {
            get
            {
                var flags = this.get_common_flags();
                if ((flags & IVisio.VisGetSetArgs.visSetFormulas) > 0)
                {
                    flags = (IVisio.VisGetSetArgs)((short)flags | (short)IVisio.VisGetSetArgs.visSetUniversalSyntax);
                }
                return flags;
            }
        }

        protected IVisio.VisGetSetArgs FormulaFlags
        {
            get
            {
                var common_flags = this.get_common_flags();
                var formula_flags = (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;
                var combined_flags = (short)common_flags | formula_flags;
                return (IVisio.VisGetSetArgs)combined_flags;
            }
        }

        private IVisio.VisGetSetArgs get_common_flags()
        {
            IVisio.VisGetSetArgs f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            IVisio.VisGetSetArgs f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = (short)f_bg | (short)f_tc;
            return (IVisio.VisGetSetArgs)flags;
        }

        public static object[] build_formulas(IList<FormulaLiteral> formulas2)
        {
            var formulas = new object[formulas2.Count];
            int i = 0;
            foreach (var rec in formulas2)
            {
                formulas[i] = rec.Value;
                i++;
            }
            return formulas;
        }

        public static void build_results(IList<ResultValue> formulas2, out object[] unitcodes, out object[] results)
        {
            unitcodes = new object[formulas2.Count];
            results = new object[formulas2.Count];
            int i = 0;
            foreach (var update in formulas2)
            {
                unitcodes[i] = update.UnitCode;
                if (update.ResultType == ResultType.ResultNumeric)
                {
                    results[i] = update.ResultNumeric;
                }
                else if (update.ResultType == ResultType.ResultString)
                {
                    results[i] = update.ResultString;
                }
                else
                {
                    throw new AutomationException("Unhandled update type");
                }
                i++;
            }
        }

        public abstract void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface);

        public void Execute(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            this.Commit(surface);
        }
        public void Execute(IVisio.Shape shape)
        {
            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(shape);
            this.Commit(surface);                
        }

        public void Execute(IVisio.Page shape)
        {
            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(shape);
            this.Commit(surface);
        }

        public void Execute(IVisio.Master shape)
        {
            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(shape);
            this.Commit(surface);
        }

        public int Count
        {
            get { return this.ValueItems.Count; }
        }

    }
}
