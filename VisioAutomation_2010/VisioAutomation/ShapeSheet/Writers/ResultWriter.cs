using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class ResultWriter : WriterBaseEx<ResultValue>
    {
        public ResultWriter() : base()
        {
        }

        public void SetResult(SRC src, string value, IVisio.VisUnitCodes unitcode)
        {
            this.SRC_StreamItems.Add(src);
            var value_item = new ResultValue(value,unitcode);
            this.SRC_ValueItems.Add( value_item);
        }

        public void SetResult(SRC src, double value, IVisio.VisUnitCodes unitcode)
        {
            this.SRC_StreamItems.Add(src);
            var value_item = new ResultValue(value, unitcode);
            this.SRC_ValueItems.Add(value_item);
        }

        public void SetResult(SIDSRC sidsrc, double value, IVisio.VisUnitCodes unitcode)
        {
            var v = new ResultValue(value, unitcode);
            this.SetResult(sidsrc, v);
        }

        public void SetResult(SIDSRC sidsrc, string value, IVisio.VisUnitCodes unitcode)
        {
            var v = new ResultValue(value, unitcode);
            this.SetResult(sidsrc, v);
        }

        public void SetResult(SIDSRC sidsrc, ResultValue v)
        {
            this.SIDSRC_StreamItems.Add(sidsrc);
            this.SIDSRC_ValueItems.Add(v);
        }

        protected override void _commit_to_surface(ShapeSheetSurface surface)
        {
            this.SRC_commit_to_surface(surface);
            this.SIDSRC_commit_to_surface(surface);
        }

        private void SIDSRC_commit_to_surface(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.SIDSRC_ValueItems.Count < 1)
            {
                return;
            }

            var stream = SIDSRC.ToStream(this.SIDSRC_StreamItems);

            var unitcodes = WriterHelper.build_results_arrays_unitcode(this.SIDSRC_ValueItems);
            var results = WriterHelper.build_results_arrays_results(this.SIDSRC_ValueItems);
            var flags = this.ComputeGetResultFlags(this.SIDSRC_ValueItems[0].ResultType);

            surface.SetResults(stream, unitcodes, results, (short)flags);
        }

        private void SRC_commit_to_surface(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.SRC_ValueItems.Count < 1)
            {
                return;
            }

            var stream = SRC.ToStream(this.SRC_StreamItems);

            var unitcodes = WriterHelper.build_results_arrays_unitcode(this.SRC_ValueItems);
            var results = WriterHelper.build_results_arrays_results(this.SRC_ValueItems);
            var flags = this.ComputeGetResultFlags(this.SRC_ValueItems[0].ResultType);
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }
}