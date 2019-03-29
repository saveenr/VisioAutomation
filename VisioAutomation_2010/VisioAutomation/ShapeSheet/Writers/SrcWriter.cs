using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SrcWriter : WriterBase
    {


        public SrcWriter() : base(CellCoordinateType.Src)
        {
        }


        public void CommitFormulas(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            this._commit_formulas(surface);
        }

        public void CommitFormulas(IVisio.Page page)
        {
            var surface = new SurfaceTarget(page);
            this._commit_formulas(surface);
        }

        public void CommitResults(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            this._commit_results(surface);
        }

        public void CommitResults(IVisio.Page page)
        {
            var surface = new SurfaceTarget(page);
            this._commit_results(surface);
        }


        public void SetValue(Src src, CellValueLiteral formula)
        {
            this.__SetValueIgnoreNull(src, formula);
        }

        public void SetValues(CellGroups.CellGroup cgb, short row)
        {
            foreach (var pair in cgb.SrcValuePairs_NewRow(row))
            {
                this.SetValue(pair.Src, pair.Value);
            }
        }

        public void SetValues(CellGroups.CellGroup cgb)
        {
            foreach (var pair in cgb.SrcValuePairs)
            {
                this.SetValue(pair.Src, pair.Value);
            }
        }

        private void __SetValueIgnoreNull(Src src, CellValueLiteral formula)
        {
            if (this._records == null)
            {
                this._records = new WriteRecordList(CellCoordinateType.Src);
            }

            if (formula.HasValue)
            {
                this._records.Add(src, formula.Value);
            }
        }

        private void _commit_formulas(SurfaceTarget surface)
        {
            if ((this._records == null || this._records.Count < 1))
            {
                return;
            }

            var stream = this._records.BuildSrcStream();
            var formulas = this._records.BuildValuesArray();

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetFormulaFlags();

            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }


        private void _commit_results(SurfaceTarget surface)
        {
            if (this._records == null || this._records.Count < 1)
            {
                return;
            }

            var stream = this._records.BuildSrcStream();
            var results = this._records.BuildValuesArray();
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