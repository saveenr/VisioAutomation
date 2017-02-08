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

        public override void Commit(ShapeSheetSurface surface)
        {
            this.SRC_commit_to_surface(surface);
            this.SIDSRC_commit_to_surface(surface);
        }

        private void SIDSRC_commit_to_surface(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.SIDSRCCount < 1)
            {
                return;
            }

            var stream = this.GetSIDSRCStream();
            var unitcodes = WriterHelper.build_unitcode_array(this.SIDSRC_Values);
            var results = WriterHelper.build_results_array(this.SIDSRC_Values);
            var flags = this.ComputeGetResultFlags(this.SIDSRC_Values[0].ResultType);

            surface.SetResults(stream, unitcodes, results, (short)flags);
        }

        private void SRC_commit_to_surface(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.SRCCount < 1)
            {
                return;
            }

            var stream = this.GetSRCStream();
            var unitcodes = WriterHelper.build_unitcode_array(this.SRC_Values);
            var results = WriterHelper.build_results_array(this.SRC_Values);
            var flags = this.ComputeGetResultFlags(this.SRC_Values[0].ResultType);
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }
}