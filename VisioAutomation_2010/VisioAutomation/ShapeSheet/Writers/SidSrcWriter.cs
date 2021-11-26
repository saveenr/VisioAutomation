 using System;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SidSrcWriter : WriterBase
    {
        public SidSrcWriter() : base(StreamType.SidSrc)
        {
        }

        public void Commit(IVisio.Page page, Core.CellValueType type)
        {
            var surface = new Core.SurfaceTarget(page);
            this.Commit(surface, type);
        }
        public void SetValue(short id, Core.Src src, Core.CellValue formula)
        {
            var sidsrc = new Core.SidSrc(id, src);
            this.__SetValueIgnoreNull(sidsrc, formula);
        }

        public void SetValue(Core.SidSrc sidsrc, Core.CellValue formula)
        {
            this.__SetValueIgnoreNull(sidsrc, formula);
        }

        public void SetValues(short id, CellGroups.CellGroup cellgroup, short row)
        {
            var pairs = cellgroup.GetSidSrcValuePairs_NewRow(id, row);
            foreach (var pair in pairs)
            {
                this.SetValue(pair.ShapeID, pair.Src, pair.Value);
            }
        }

        public void Commit(IVisio.Page page, object formula)
        {
            throw new NotImplementedException();
        }

        public void SetValues(short id, CellGroups.CellGroup cellgroup)
        {
            foreach (var pair in cellgroup.GetSrcValuePairs())
            {
                this.SetValue(id, pair.Src, pair.Value);
            }
        }

        private void __SetValueIgnoreNull(Core.SidSrc sidsrc, Core.CellValue formula)
        {
            if (this._records == null)
            {
                this._records = new WriteRecordList(StreamType.SidSrc);
            }

            if (formula.HasValue)
            {
                this._records.Add(sidsrc, formula.Value);
            }
        }

        public void CommitFormulas(Core.SurfaceTarget surface)
        {
            if ((this._records == null || this._records.Count < 1))
            {
                return;
            }

            var stream = this._records.BuildStreamArray(StreamType.SidSrc);
            var formulas = this._records.BuildValuesArray();

            if (stream.Array.Length == 0)
            {
                throw new Exceptions.InternalAssertionException();
            }

            var flags = this._compute_setformula_flags();

            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        public void Commit(Core.SurfaceTarget surface, Core.CellValueType type)
        {
            if ((this._records == null || this._records.Count < 1))
            {
                return;
            }

            var stream = this._records.BuildStreamArray(StreamType.SidSrc);
            var items = this._records.BuildValuesArray();

            if (stream.Array.Length == 0)
            {
                throw new Exceptions.InternalAssertionException();
            }

            if (type == Core.CellValueType.Formula)
            {
                var flags = this._compute_setformula_flags();
                int c = surface.SetFormulas(stream, items, (short)flags);
            }
            else
            {
                const object[] unitcodes = null;
                var flags = this._compute_setresults_flags();
                surface.SetResults(stream, unitcodes, items, (short)flags);
            }
        }
    }
}