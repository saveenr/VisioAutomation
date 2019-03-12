using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SidSrcWriter : WriterBase
    {

        private WriteCache<SidSrc> _formulaRecords;
        private WriteCache<SidSrc> _resultRecords;

        public SidSrcWriter()
        {
        }

        public void Clear()
        {
            _formulaRecords?.Clear();
            _resultRecords?.Clear();
        }

        public void Commit(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            this.Commit(surface);
        }

        public void Commit(IVisio.Page page)
        {
            var surface = new SurfaceTarget(page);
            this.Commit(surface);
        }

        public void Commit(VisioAutomation.SurfaceTarget surface)
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
            if (this._formulaRecords == null)
            {
                this._formulaRecords = new WriteCache<SidSrc>();
            }

            if (formula.HasValue)
            {
                this._formulaRecords.Add(sidsrc, formula.Value);
            }
        }

        private VisioAutomation.ShapeSheet.Streams.StreamArray buildstream_sidsrc(WriteCache<SidSrc> wcs)
        {
            var builder = new VisioAutomation.ShapeSheet.Streams.SidSrcStreamArrayBuilder(wcs.Count);
            builder.AddRange(wcs.EnumCoords());
            return builder.ToStreamArray();
        }

        private void CommitFormulas(SurfaceTarget surface)
        {
            if ((this._formulaRecords == null || this._formulaRecords.Count < 1))
            {
                return;
            }

            var stream = this.buildstream_sidsrc(this._formulaRecords);
            var formulas = this._formulaRecords.BuildValues();

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
            if (this._resultRecords == null)
            {
                this._resultRecords = new WriteCache<SidSrc>();
            }

            this._resultRecords.Add(sidsrc, result.Value);
        }

        private void CommitResults(SurfaceTarget surface)
        {
            if ((this._resultRecords == null || this._resultRecords.Count < 1))
            {
                return;
            }

            var stream = this.buildstream_sidsrc(this._resultRecords);
            var results = this._resultRecords.BuildValues();
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