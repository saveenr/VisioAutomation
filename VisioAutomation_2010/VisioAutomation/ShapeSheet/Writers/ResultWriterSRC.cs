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


        public void SetResult(SRC streamitem, string value, Microsoft.Office.Interop.Visio.VisUnitCodes unitcode)
        {
            this.StreamItems.Add(streamitem);
            this.ValueItems.Add( new ResultValue(value,unitcode));
        }

        public void SetResult(SRC streamitem, double value, Microsoft.Office.Interop.Visio.VisUnitCodes unitcode)
        {
            this.StreamItems.Add(streamitem);
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
            var flags = this.GetResultFlags();
            if (this.ValueItems[0].ResultType == ResultType.ResultString)
            {
                flags |= Microsoft.Office.Interop.Visio.VisGetSetArgs.visGetStrings;
            }
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }

    }
}