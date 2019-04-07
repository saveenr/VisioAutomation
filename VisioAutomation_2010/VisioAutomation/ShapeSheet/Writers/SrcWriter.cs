using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SrcWriter : WriterBase
    {


        public SrcWriter() : base(StreamType.Src)
        {
        }


        public void Commit(IVisio.Shape shape, CellValueType type)
        {
            var surface = new SurfaceTarget(shape);
            this._commit(surface, type);
        }

        public void Commit(IVisio.Page page, CellValueType type)
        {
            var surface = new SurfaceTarget(page);
            this._commit(surface, type);
        }

        public void SetValue(Src src, CellValue formula)
        {
            this.__set_value_ignore_null(src, formula);
        }

        public void SetValues(CellGroups.CellGroup cellgroup, short row)
        {
            foreach (var pair in cellgroup.GetSrcValuePairs_NewRow(row))
            {
                this.SetValue(pair.Src, pair.Value);
            }
        }

        public void SetValues(CellGroups.CellGroup cellgroup)
        {
            foreach (var pair in cellgroup.GetSrcValuePairs())
            {
                this.SetValue(pair.Src, pair.Value);
            }
        }

        private void __set_value_ignore_null(Src src, CellValue formula)
        {
            if (this._records == null)
            {
                this._records = new WriteRecordList(StreamType.Src);
            }

            if (formula.HasValue)
            {
                this._records.Add(src, formula.Value);
            }
        }

        private void _commit(SurfaceTarget surface, CellValueType type)
        {
            if (this._records == null || this._records.Count < 1)
            {
                return;
            }

            var stream = this._records.BuildStreamArray(StreamType.Src);

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var values = this._records.BuildValuesArray();

            if (type == CellValueType.Formula)
            {
                var flags = this._compute_setformula_flags();
                int c = surface.SetFormulas(stream, values, (short)flags);

            }
            else
            {
                const object[] unitcodes = null;
                var flags = this._compute_setresults_flags();
                surface.SetResults(stream, unitcodes, values, (short)flags);
            }
        }
    }
}