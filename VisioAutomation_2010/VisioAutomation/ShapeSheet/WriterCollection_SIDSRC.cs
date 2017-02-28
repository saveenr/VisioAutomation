using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    class WriterCollection_SIDSRC
    {
        private SIDSRCStreamBuilder StreamBuilder;
        private ShapeSheetObjectArrayBuilder<string> ValuesBuilder;
        private ShapeSheetObjectArrayBuilder<Microsoft.Office.Interop.Visio.VisUnitCodes> UnitCodesBuilder;

        private bool include_unit_codes;

        public WriterCollection_SIDSRC(bool include_unit_codes)
        {
            this.include_unit_codes = include_unit_codes;
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

        public void Add(SIDSRC sidsrc, string value)
        {
            if (this.include_unit_codes)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            this.StreamBuilder.Add(sidsrc);
            this.ValuesBuilder.Add(value);
        }

        public void Add(SIDSRC sidsrc, string value, Microsoft.Office.Interop.Visio.VisUnitCodes unit_code)
        {
            if (!this.include_unit_codes)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            this.StreamBuilder.Add(sidsrc);
            this.ValuesBuilder.Add(value);
            this.UnitCodesBuilder.Add(unit_code);
        }


        public short[] BuildStream()
        {
            return this.StreamBuilder.ToStream();
        }

        public object[] BuildValues()
        {
            return this.ValuesBuilder.ToObjectArray();
        }

        public object[] BuildUnitCodes()
        {
            return this.UnitCodesBuilder.ToObjectArray();
        }

        public int Count => this.StreamBuilder.Count();

    }
}