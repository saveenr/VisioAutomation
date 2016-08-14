using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class ResultWriterSIDSRC : WriterBase<VisioAutomation.ShapeSheet.SIDSRC, ResultValue>
    {

        public ResultWriterSIDSRC() : base()
        {
        }

        public ResultWriterSIDSRC(int capacity) : base(capacity)
        {
        }
        
        public void SetResult(short shapeid, SRC src, double value, IVisio.VisUnitCodes unitcode)
        {
            var streamitem = new SIDSRC(shapeid, src);
            var v = new ResultValue(value,unitcode);
            this.SetResult(streamitem, v);
        }

        public void SetResult(short shapeid, SRC src, string value, IVisio.VisUnitCodes unitcode)
        {
            var streamitem = new SIDSRC(shapeid, src);
            var v = new ResultValue(value, unitcode);
            this.SetResult(streamitem, v);
        }

        public void SetResult(SIDSRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            var v = new ResultValue(value, unitcode);
            this.SetResult(streamitem, v);
        }

        public void SetResult(SIDSRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            var v = new ResultValue(value, unitcode);
            this.SetResult(streamitem, v);
        }

        public void SetResult(SIDSRC streamitem, ResultValue v)
        {
            this.StreamItems.Add(streamitem);
            this.ValueItems.Add(v);
        }

        protected override void _commit_to_surface(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.ValueItems.Count < 1)
            {
                return;
            }

            var stream = SIDSRC.ToStream(this.StreamItems);

            var unitcodes = WriterHelper.build_results_arrays_unitcode(this.ValueItems);
            var results = WriterHelper.build_results_arrays_results(this.ValueItems);

            var flags = this.GetResultFlags();
            if (this.ValueItems[0].ResultType == ResultType.ResultString)
            {
                flags |= IVisio.VisGetSetArgs.visGetStrings;
            }
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }
}