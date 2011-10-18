using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Collections;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Update
{
    public class UpdateBase<T> : IEnumerable<UpdateRecord<T>>
        where T : struct
    {
        private List<UpdateRecord<T>> items;
        public int ResultCount { get; private set; }
        public int FormulaCount { get; private set; }
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        protected UpdateBase()
        {
            this.items = new List<UpdateRecord<T>>();
        }

        protected UpdateBase(int capacity)
        {
            this.items = new List<UpdateRecord<T>>(capacity);
        }

        public IVisio.VisGetSetArgs ResultFlags
        {
            get { return get_common_flags(); }
        }

        public IVisio.VisGetSetArgs FormulaFlags
        {
            get
            {
                var common_flags = get_common_flags();
                var formula_flags = (short) IVisio.VisGetSetArgs.visSetUniversalSyntax;
                var combined_flags = (short) common_flags | formula_flags;
                return (IVisio.VisGetSetArgs) combined_flags;
            }
        }

        private IVisio.VisGetSetArgs get_common_flags()
        {
            IVisio.VisGetSetArgs f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            IVisio.VisGetSetArgs f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = (short) f_bg | (short) f_tc;
            return (IVisio.VisGetSetArgs) flags;
        }

        public void SetFormula(T streamitem, FormulaLiteral literal)
        {
            ShapeSheetHelper.CheckFormulaIsNotNull(literal.Value);
            var rec = new UpdateRecord<T>(streamitem, literal.Value);
            this.items.Add(rec);
            this.FormulaCount++;
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
            var rec = new UpdateRecord<T>(streamitem, value, unitcode);
            this.items.Add(rec);
            this.ResultCount++;
        }

        public IEnumerator<UpdateRecord<T>> GetEnumerator()
        {
            foreach (var i in this.items)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator() // Explicit implementation
        {
            // keeps it hidden.
            return GetEnumerator();
        }

        public IEnumerable<UpdateRecord<T>> ResultRecords
        {
            get { return this.items.Where(i => i.UpdateType == VA.ShapeSheet.Update.UpdateType.Result); }
        }

        public IEnumerable<UpdateRecord<T>> FormulaRecords
        {
            get { return this.items.Where(i => i.UpdateType == VA.ShapeSheet.Update.UpdateType.Formula); }
        }

        public string[] GetFormulasArray()
        {
            var a = new string[this.FormulaCount];
            int i = 0;
            foreach (var rec in this.FormulaRecords)
            {
                a[i] = rec.Formula;
                i++;
            }
            return a;
        }

        public double[] GetResultsArray()
        {
            var a = new double[this.ResultCount];
            int i = 0;
            foreach (var rec in this.ResultRecords)
            {
                a[i] = rec.Result;
                i++;
            }
            return a;
        }

        public IVisio.VisUnitCodes[] GetUnitCodesArray()
        {
            var a = new IVisio.VisUnitCodes[this.ResultCount];
            int i = 0;
            foreach (var rec in this.ResultRecords)
            {
                a[i] = rec.UnitCode;
                i++;
            }
            return a;
        }

    }
}