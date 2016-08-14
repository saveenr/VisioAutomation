namespace VisioAutomation.ShapeSheet.Writers
{
    public class SRCResultWriter : WriterBase<VisioAutomation.ShapeSheet.SRC, ResultValue>
    {
        public SRCResultWriter() : base()
        {
        }

        public SRCResultWriter(int capacity) : base(capacity)
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

        public override void Commit(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.ValueItems.Count < 1)
            {
                return;
            }

            var stream = SRC.ToStream(this.StreamItems);

            object[] unitcodes;
            object[] results;

            WriterBase<SIDSRC, ResultValue>.build_results(this.ValueItems, out unitcodes, out results);
            var flags = this.ResultFlags;
            if (this.ValueItems[0].ResultType == ResultType.ResultString)
            {
                flags |= Microsoft.Office.Interop.Visio.VisGetSetArgs.visGetStrings;
            }
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }

    }
}