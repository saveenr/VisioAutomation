using VisioAutomation.ShapeSheet.Internal;

namespace VisioAutomation.ShapeSheet
{
    public class ShapeSheetWriterSidSrc : ShapeSheetWriterBase
    {

        private WriterCollection<SidSrc> FormulaRecords;
        private WriterCollection<SidSrc> ResultRecords;

        public ShapeSheetWriterSidSrc()
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
            this.CommitFormulas(surface);
            this.CommitResults(surface);
        }

        public void SetFormula(short id, Src src, CellValueLiteral formula)
        {
            var sidsrc = new SidSrc(id, src);
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        public void SetFormula(SidSrc sidsrc, CellValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        private void __SetFormulaIgnoreNull(SidSrc sidsrc, CellValueLiteral formula)
        {
            if (this.FormulaRecords == null)
            {
                this.FormulaRecords = new WriterCollection<SidSrc>();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords.Add(sidsrc, formula.Value);
            }
        }

        private VisioAutomation.ShapeSheet.Streams.StreamArray buildstream_sidsrc(WriterCollection<SidSrc> wcs)
        {
            var builder = new VisioAutomation.ShapeSheet.Streams.FixedSidSrcStreamBuilder(wcs.Count);
            builder.AddRange(wcs.EnumCoords());
            return builder.ToStream();
        }

        private void CommitFormulas(ShapeSheetSurface surface)
        {
            if ((this.FormulaRecords == null || this.FormulaRecords.Count < 1))
            {
                return;
            }

            var stream = this.buildstream_sidsrc(this.FormulaRecords);
            var formulas = this.FormulaRecords.BuildValues();

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetFormulaFlags();

            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        public void SetResult(short id, Src src, CellValueLiteral result)
        {
            var sidsrc = new SidSrc(id, src);
            this.SetResult(sidsrc, result.Value);
        }

        public void SetResult(SidSrc sidsrc, CellValueLiteral result)
        {
            if (this.ResultRecords == null)
            {
                this.ResultRecords = new WriterCollection<SidSrc>();
            }

            this.ResultRecords.Add(sidsrc, result.Value);
        }

        private void CommitResults(ShapeSheetSurface surface)
        {
            if ((this.ResultRecords == null || this.ResultRecords.Count < 1))
            {
                return;
            }

            var stream = this.buildstream_sidsrc(this.ResultRecords);
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