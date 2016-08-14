namespace VisioAutomation.ShapeSheet.Writers
{
    public class SRCResultWriter : SRCWriter
    {
        public SRCResultWriter() : base()
        {
        }

        public SRCResultWriter(int capacity) : base(capacity)
        {
        }


        public void SetResult(SRC streamitem, string value, Microsoft.Office.Interop.Visio.VisUnitCodes unitcode)
        {
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(SRC streamitem, double value, Microsoft.Office.Interop.Visio.VisUnitCodes unitcode)
        {
            this._SetResult(streamitem, value, unitcode);
        }

        protected void _SetResult(SRC streamitem, double value, Microsoft.Office.Interop.Visio.VisUnitCodes unitcode)
        {
            var rec = new WriterRecord<SRC>(streamitem, value, unitcode);
            this._add_update(rec);
        }

        protected void _SetResult(SRC streamitem, string value, Microsoft.Office.Interop.Visio.VisUnitCodes unitcode)
        {
            var rec = new WriterRecord<SRC>(streamitem, value, unitcode);
            this._add_update(rec);
        }

    }
}