using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class ResultWriterSRC : WriterBase<VisioAutomation.ShapeSheet.SRC, ResultValue>
    {
        public ResultWriterSRC() : base()
        {
        }

        public ResultWriterSRC(int capacity) : base(capacity)
        {
        }


        public void SetResult(SRC src, string value, IVisio.VisUnitCodes unitcode)
        {
            this.StreamItems.Add(src);
            this.ValueItems.Add( new ResultValue(value,unitcode));
        }

        public void SetResult(SRC src, double value, IVisio.VisUnitCodes unitcode)
        {
            this.StreamItems.Add(src);
            this.ValueItems.Add(new ResultValue(value, unitcode));
        }

        protected override void _commit_to_surface(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.ValueItems.Count < 1)
            {
                return;
            }

            var stream = SRC.ToStream(this.StreamItems);

            var unitcodes = WriterHelper.build_results_arrays_unitcode(this.ValueItems);
            var results = WriterHelper.build_results_arrays_results(this.ValueItems);
            var flags = this.ComputeGetResultFlags(this.ValueItems[0].ResultType);
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }

    }
}