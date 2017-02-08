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

        public static object[] build_unitcode_array(IList<ResultValue> result_values)
        {
            var unitcodes = new object[result_values.Count];
            int i = 0;
            foreach (var result_value in result_values)
            {
                unitcodes[i] = result_value.UnitCode;
                i++;
            }
            return unitcodes;
        }

        public static object[] build_results_array(IList<ResultValue> result_values)
        {
            var results = new object[result_values.Count];
            int i = 0;
            foreach (var result_value in result_values)
            {
                if (result_value.ResultType == ResultType.ResultNumeric)
                {
                    results[i] = result_value.ValueNumeric;
                }
                else if (result_value.ResultType == ResultType.ResultString)
                {
                    results[i] = result_value.ValueString;
                }
                else
                {
                    string msg = string.Format("Unsupported {0}.{1} \"{2}\"", nameof(result_value),nameof(result_value.ResultType),result_value.ResultType);
                    throw new System.ArgumentOutOfRangeException(msg);
                }
                i++;
            }

            return results;
        }
    }
}