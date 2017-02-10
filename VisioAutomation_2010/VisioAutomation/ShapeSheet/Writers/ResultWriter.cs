using System.Collections.Generic;
using System.Linq;
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

        protected override void CommitRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            var records = this.GetRecords(coord_type);
            var count = records.Count();

            if (count == 0)
            {
                return;
            }

            int chunksize = coord_type == CoordType.SIDSRC ? 4 : 3;

            var stream = new short[count * chunksize];
            var results = new object[count];
            var unitcodes = new object[count];

            int streampos = 0;
            int resultspos = 0;
            int unitcodespos = 0;

            foreach (var rec in records)
            {
                // fill stream
                if (coord_type == CoordType.SRC)
                {
                    var src = rec.SRC;
                    stream[streampos++] = src.Section;
                    stream[streampos++] = src.Row;
                    stream[streampos++] = src.Cell;
                }
                else
                {
                    var sidsrc = rec.SIDSRC;
                    stream[streampos++] = sidsrc.ShapeID;
                    stream[streampos++] = sidsrc.Section;
                    stream[streampos++] = sidsrc.Row;
                    stream[streampos++] = sidsrc.Cell;
                }

                // fill results
                if (rec.Value.ResultType == ResultType.ResultNumeric)
                {
                    results[resultspos++] = rec.Value.ValueNumeric;
                }
                else if (rec.Value.ResultType == ResultType.ResultString)
                {
                    results[resultspos++] = rec.Value.ValueString;
                }

                // fill unit codes
                unitcodes[unitcodespos] = rec.Value.UnitCode;

            }

            var flags = this.ComputeGetResultFlags(records.First().Value.ResultType);
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }
}