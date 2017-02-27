namespace VisioAutomation.ShapeSheet
{
    class WriterCollection_SIDSRC
    {
        public SIDSRCStreamBuilder StreamBuilder;
        public ShapeSheetObjectArrayBuilder<string> ValuesBuilder;
        public ShapeSheetObjectArrayBuilder<Microsoft.Office.Interop.Visio.VisUnitCodes> UnitCodesBuilder;

        public WriterCollection_SIDSRC(bool include_unit_codes)
        {
            this.StreamBuilder = new SIDSRCStreamBuilder();
            this.ValuesBuilder = new ShapeSheetObjectArrayBuilder<string>();
            if (include_unit_codes)
            {
                this.UnitCodesBuilder = new ShapeSheetObjectArrayBuilder<Microsoft.Office.Interop.Visio.VisUnitCodes>();
            }
        }

        public void Clear()
        {
            this.StreamBuilder.Clear();
            this.ValuesBuilder.Clear();
            UnitCodesBuilder?.Clear();
        }
    }
}