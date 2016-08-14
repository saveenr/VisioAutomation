using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SIDSRCResultWriter : WriterBase<VisioAutomation.ShapeSheet.SIDSRC, ResultValue>
    {

        public SIDSRCResultWriter() : base()
        {
        }

        public SIDSRCResultWriter(int capacity) : base(capacity)
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

        public override void Commit(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.ValueItems.Count < 1)
            {
                return;
            }

            var stream = SIDSRC.ToStream(this.StreamItems);

            object[] unitcodes;
            object[] results;

            WriterBase<SIDSRC,ResultValue>.build_results(this.ValueItems,out unitcodes, out results);
            var flags = this.ResultFlags;
            if (this.ValueItems[0].ResultType == ResultType.ResultString)
            {
                flags |= IVisio.VisGetSetArgs.visGetStrings;
            }
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }
}