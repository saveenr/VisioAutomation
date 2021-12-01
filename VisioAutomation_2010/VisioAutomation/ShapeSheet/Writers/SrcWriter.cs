using VisioAutomation.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SrcWriter : WriterBase
    {
        public SrcWriter() : base(VisioAutomation.ShapeSheet.Streams.StreamType.Src)
        {
        }


        public void Commit(IVisio.Shape shape, Core.CellValueType type)
        {
            var visobjtarget = new VisioObjectTarget(shape);
            this._commit(visobjtarget, type);
        }

        public void Commit(IVisio.Master master, Core.CellValueType type)
        {
            var visobjtarget = new VisioObjectTarget(master);

            this._commit(visobjtarget, type);
        }

        public void Commit(IVisio.Page page, Core.CellValueType type)
        {
            var visobjtarget = new VisioObjectTarget(page);
            this._commit(visobjtarget, type);
        }

        public void SetValue(Core.Src src, Core.CellValue formula)
        {
            this.__set_value_ignore_null(src, formula);
        }

        public void SetValues(CellGroups.CellGroup cellgroup, short row)
        {
            foreach (var pair in cellgroup.GetSrcValuesWithNewRow(row))
            {
                this.SetValue(pair.Src, pair.Value);
            }
        }

        public void SetValues(CellGroups.CellGroup cellgroup)
        {
            foreach (var pair in cellgroup.GetSrcValues())
            {
                this.SetValue(pair.Src, pair.Value);
            }
        }

        private void __set_value_ignore_null(Core.Src src, Core.CellValue formula)
        {
            if (this._records == null)
            {
                this._records = new WriteRecordList(VisioAutomation.ShapeSheet.Streams.StreamType.Src);
            }

            if (formula.HasValue)
            {
                this._records.Add(src, formula.Value);
            }
        }

        private void _commit(VisioObjectTarget visobjtarget, Core.CellValueType type)
        {
            if (this._records == null || this._records.Count < 1)
            {
                return;
            }

            var stream = this._records.BuildStreamArray(VisioAutomation.ShapeSheet.Streams.StreamType.Src);

            if (stream.Array.Length == 0)
            {
                throw new Exceptions.InternalAssertionException();
            }

            var values = this._records.BuildValuesArray();

            if (type == Core.CellValueType.Formula)
            {
                var flags = this._compute_setformula_flags();
                var c = visobjtarget.SetFormulas(stream, values, (short) flags);
            }
            else
            {
                const object[] unitcodes = null;
                var flags = this._compute_setresults_flags();
                var c = visobjtarget.SetResults(stream, unitcodes, values, (short) flags);
            }
        }
    }
}