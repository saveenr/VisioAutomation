using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Writers
{
    internal static class WriterHelper
    {
        public static object[] build_formulas_array(IList<FormulaLiteral> formulas)
        {
            var result = new object[formulas.Count];
            int i = 0;
            foreach (var rec in formulas)
            {
                result[i] = rec.Value;
                i++;
            }
            return result;
        }

        public static object[] build_results_arrays_unitcode(IList<ResultValue> formulas2)
        {
            var unitcodes = new object[formulas2.Count];
            int i = 0;
            foreach (var update in formulas2)
            {
                unitcodes[i] = update.UnitCode;
                i++;
            }
            return unitcodes;
        }

        public static object[] build_results_arrays_results(IList<ResultValue> formulas2)
        {
            var results = new object[formulas2.Count];
            int i = 0;
            foreach (var update in formulas2)
            {
                if (update.ResultType == ResultType.ResultNumeric)
                {
                    results[i] = update.ValueNumeric;
                }
                else if (update.ResultType == ResultType.ResultString)
                {
                    results[i] = update.ValueString;
                }
                else
                {
                    string msg = string.Format("Unsupported {0}.{1} \"{2}\"", nameof(update),nameof(update.ResultType),update.ResultType);
                    throw new System.ArgumentOutOfRangeException(msg);
                }
                i++;
            }

            return results;
        }
    }
}