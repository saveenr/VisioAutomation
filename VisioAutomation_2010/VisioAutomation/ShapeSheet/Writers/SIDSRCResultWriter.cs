using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SIDSRCResultWriter : SIDSRCWriter
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
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(SIDSRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(short shapeid, SRC src, string value, IVisio.VisUnitCodes unitcode)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(SIDSRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(streamitem, value, unitcode);
        }


        protected void _SetResult(SIDSRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            var rec = new WriterRecord<SIDSRC>(streamitem, value, unitcode);
            this._add_update(rec);
        }

        protected void _SetResult(SIDSRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            var rec = new WriterRecord<SIDSRC>(streamitem, value, unitcode);
            this._add_update(rec);
        }
    }
}