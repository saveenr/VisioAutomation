using VisioAutomation.ShapeSheet.Internal;

namespace VisioAutomation.ShapeSheet
{
    public class ShapeSheetWriterSrc : ShapeSheetWriterBase
    {
        private WriterCollection<Src> FormulaRecords;
        private WriterCollection<Src> ResultRecords;

        public ShapeSheetWriterSrc()
        {
        }

        public void Clear()
        {
            FormulaRecords?.Clear();
            ResultRecords?.Clear();
        }

        public void Commit(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var surface = new ShapeSheetSurface(shape);
            this.Commit(surface);
        }

        public void Commit(Microsoft.Office.Interop.Visio.Page page)
        {
            var surface = new ShapeSheetSurface(page);
            this.Commit(surface);
        }

        public void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            this.CommitFormulaRecordsByType(surface);
            this.CommitResultRecordsByType(surface);
        }

        public void SetFormula(Src src, CellValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(src, formula);
        }
        
        private void __SetFormulaIgnoreNull(Src src, CellValueLiteral formula)
        {
            if (this.FormulaRecords == null)
            {
                this.FormulaRecords = new WriterCollection<Src>();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords.Add(src, formula.Value);
            }
        }

        private VisioAutomation.ShapeSheet.Streams.StreamArray buildstream_src(WriterCollection<Src> wcs)
        {
            var builder = new VisioAutomation.ShapeSheet.Streams.FixedSrcStreamBuilder(wcs.Count);
            builder.AddRange(wcs.EnumCoords());
            return builder.ToStream();
        }

        private void CommitFormulaRecordsByType(ShapeSheetSurface surface)
        {
            if ((this.FormulaRecords == null || this.FormulaRecords.Count < 1))
            {
                return;
            }

            var stream = this.buildstream_src(this.FormulaRecords);
            var formulas = this.FormulaRecords.BuildValues();

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetFormulaFlags();

            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        public void SetResult(Src src, CellValueLiteral result)
        {
            if (this.ResultRecords == null)
            {
                this.ResultRecords = new WriterCollection<Src>();
            }

            this.ResultRecords.Add(src, result.Value);
        }

        private void CommitResultRecordsByType(ShapeSheetSurface surface)
        {
            if (this.ResultRecords == null || this.ResultRecords.Count < 1)
            {
                return;
            }

            var stream = this.buildstream_src(this.ResultRecords);
            var results = this.ResultRecords.BuildValues();
            const object[] unitcodes = null;

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetResultFlags();
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }
}