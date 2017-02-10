using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class ResultWriter : WriterBase<ResultValue>
    {
        public ResultWriter() : base()
        {
        }

        public void SetResult(SRC src, string value, IVisio.VisUnitCodes unitcode)
        {
            var value_item = new ResultValue(value, unitcode);
            this.Add(src,value_item);
        }

        public void SetResult(SRC src, double value, IVisio.VisUnitCodes unitcode)
        {
            var value_item = new ResultValue(value, unitcode);
            this.Add(src,value_item);
        }

        public void SetResult(SIDSRC sidsrc, double value, IVisio.VisUnitCodes unitcode)
        {
            var v = new ResultValue(value, unitcode);
            this.Add(sidsrc, v);
        }

        public void SetResult(SIDSRC sidsrc, string value, IVisio.VisUnitCodes unitcode)
        {
            var v = new ResultValue(value, unitcode);
            this.Add(sidsrc, v);
        }

        protected override void CommitSIDSRC(ShapeSheetSurface surface)
        {
            var stream = this.GetSIDSRCStream();
            var unitcodes = build_unitcode_array(this.SIDSRC_Values);
            var results = build_results_array(this.SIDSRC_Values);
            var flags = this.ComputeGetResultFlags(this.SIDSRC_Values[0].ResultType);

            surface.SetResults(stream, unitcodes, results, (short)flags);
        }

        protected override void CommitSRC(ShapeSheetSurface surface)
        {
            var stream = this.GetSRCStream();
            var unitcodes = build_unitcode_array(this.SRC_Values);
            var results = build_results_array(this.SRC_Values);
            var flags = this.ComputeGetResultFlags(this.SRC_Values[0].ResultType);
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }

        private static object[] build_unitcode_array(IList<ResultValue> result_values)
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

        private static object[] build_results_array(IList<ResultValue> result_values)
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
                    string msg = string.Format("Unsupported {0}.{1} \"{2}\"", nameof(result_value), nameof(result_value.ResultType), result_value.ResultType);
                    throw new System.ArgumentOutOfRangeException(msg);
                }
                i++;
            }

            return results;
        }
    }
}