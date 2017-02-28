namespace VisioAutomation.ShapeSheet
{
    class WriterCollection_SRC
    {
        public SRCStreamBuilder StreamBuilder;
        public ShapeSheetObjectArrayBuilder<string> ValuesBuilder;
        public ShapeSheetObjectArrayBuilder<Microsoft.Office.Interop.Visio.VisUnitCodes> UnitCodesBuilder;

        private bool include_unit_codes;

        public WriterCollection_SRC(bool include_unit_codes)
        {
            this.include_unit_codes = include_unit_codes;

            this.StreamBuilder = new SRCStreamBuilder();
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


        public void Add(SRC src, string value)
        {
            if (this.include_unit_codes)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            this.StreamBuilder.Add(src);
            this.ValuesBuilder.Add(value);
        }

        public void Add(SRC src, string value, Microsoft.Office.Interop.Visio.VisUnitCodes unit_code)
        {
            if (!this.include_unit_codes)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            this.StreamBuilder.Add(src);
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