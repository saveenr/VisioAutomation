using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Collections;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Update
{
    public static class UpdateUtil
    {
        internal static IVisio.VisGetSetArgs _CheckSetResultsFlags(IVisio.VisGetSetArgs flags)
        {
            if ((flags & IVisio.VisGetSetArgs.visSetUniversalSyntax) > 0)
            {
                string msg = string.Format("visSetUniversalSyntax allowed only with visSetFormulas");
                throw new AutomationException(msg);
            }

            // force universal syntax if strings are set as formulas
            // if SetResults will fail if UniversalSyntax flag is used alone
            if ((flags & IVisio.VisGetSetArgs.visSetFormulas) > 0)
            {
                flags = (IVisio.VisGetSetArgs)((short)flags | (short)IVisio.VisGetSetArgs.visSetUniversalSyntax);
            }

            return flags;
        }
        
        internal static object[] StringsToObjectArray(IList<string> strings)
        {
            if (strings == null)
            {
                return null;
            }

            int num_items = strings.Count;
            var destination_array = new object[num_items];
            for (int i = 0; i < num_items; i++)
            {
                destination_array[i] = strings[i];
            }
            return destination_array;
        }


        internal static object[] DoublesToObjectArray(IList<double> doubles)
        {
            if (doubles == null)
            {
                return null;
            }

            int num_items = doubles.Count;
            var destination_array = new object[num_items];
            for (int i = 0; i < num_items; i++)
            {
                destination_array[i] = doubles[i];
            }
            return destination_array;
        }

        internal static short SetFormulas(
    IVisio.Page page,
    short[] stream,
    IList<string> formulas,
    short flags,
    int numitems)
        {
            if (numitems < 1)
            {
                return 0;
            }

            var formula_obj_array = UpdateUtil.StringsToObjectArray(formulas);

            // Force UniversalSyntax 
            flags |= (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;

            return page.SetFormulas(stream, formula_obj_array, flags);
        }


        internal static short SetFormulas(
    IVisio.Shape shape,
    short[] stream,
    IList<string> formulas,
    IVisio.VisGetSetArgs flags,
            int numitems)
        {
            if (formulas.Count != numitems)
            {
                string msg = string.Format("Expected {0} formulas, instead have {1}", numitems, formulas.Count);
                throw new AutomationException(msg);
            }

            if (numitems == 0)
            {
                return 0;
            }


            var formula_obj_array = UpdateUtil.StringsToObjectArray(formulas);

            // Force UniversalSyntax 
            short short_flags = (short)(((short)flags) | ((short)IVisio.VisGetSetArgs.visSetUniversalSyntax));

            return shape.SetFormulas(stream, formula_obj_array, short_flags);
        }

        internal static short SetResults(
    IVisio.Shape shape,
    short[] stream,
    IList<double> results,
    IList<IVisio.VisUnitCodes> unit_codes,
    IVisio.VisGetSetArgs flags,
            int numitems)
        {
            if (unit_codes.Count != numitems)
            {
                string msg = string.Format("Expected {0} unit_codes, instead have {1}", numitems, unit_codes.Count);
                throw new AutomationException(msg);
            }

            if (results.Count != numitems)
            {
                string msg = string.Format("Expected {0} results, instead have {1}", numitems, results.Count);
                throw new AutomationException(msg);
            }

            if (numitems < 1)
            {
                return 0;
            }

            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unit_codes);
            var results_obj_array = UpdateUtil.DoublesToObjectArray(results);

            flags = UpdateUtil._CheckSetResultsFlags(flags);

            short num_set = shape.SetResults(stream, unitcodes_obj_array, results_obj_array, (short)flags);

            return num_set;
        }

        internal static short SetResults(
    IVisio.Page page,
    short[] stream,
    IList<double> results,
    IList<IVisio.VisUnitCodes> unitcodes,
    IVisio.VisGetSetArgs flags,
            int numitems)
        {
            if (results.Count != numitems)
            {
                string msg = string.Format("Expected {0} results, instead have {1}", numitems, results.Count);
                throw new AutomationException(msg);
            }

            if (numitems == 0)
            {
                return 0;
            }

            var results_obj_array = UpdateUtil.DoublesToObjectArray(results);
            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unitcodes);

            flags = UpdateUtil._CheckSetResultsFlags(flags);

            return page.SetResults(stream, unitcodes_obj_array, results_obj_array, (short)flags);
        }

    }
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


        void CheckFormulaIsNotNull(string formula)
        {
            if (formula == null)
            {
                throw new AutomationException("Null not allowed for formula");
            }
        }

        public void SetFormula(T streamitem, FormulaLiteral literal)
        {
            this.CheckFormulaIsNotNull(literal.Value);
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