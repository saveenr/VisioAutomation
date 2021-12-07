using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Internal;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet.CellRecords;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SidSrcWriter : WriterBase
    {
        public SidSrcWriter() : base(VA.ShapeSheet.Streams.StreamType.SidSrc)
        {
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

        public void SetValues(short id, CellRecord cellrecord, short row)
        {

            var metadata = cellrecord.GetCellMetadata();
            var srcvalues = metadata.Select(i => new SidSrcValue(id, i.Src.CloneWithNewRow(row), i.Value));

            foreach (var item in srcvalues)
            {
                this.SetValue(item.ShapeID, item.Src, item.Value);
            }
        }

        public void SetValues(short id, CellRecord cellrecord)
        {
            foreach (var pair in cellrecord.GetCellMetadata())
            {
                this.SetValue(id, pair.Src, pair.Value);
            }
        }

        private void __SetValueIgnoreNull(Core.SidSrc sidsrc, Core.CellValue formula)
        {
            if (this._records == null)
            {
                this._records = new WriteRecordList(VA.ShapeSheet.Streams.StreamType.SidSrc);
            }

            if (formula.HasValue)
            {
                this._records.Add(sidsrc, formula.Value);
            }
        }


        public void Commit(IVisio.Page page, Core.CellValueType type)
        {
            var visobjtarget = new VisioObjectTarget(page);
            this._commit(visobjtarget, type);
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

        private void _commit(VisioObjectTarget visobjtarget, Core.CellValueType type)
        {
            if ((this._records == null || this._records.Count < 1))
            {
                return;
            }

            var stream = this._records.BuildStreamArray(VA.ShapeSheet.Streams.StreamType.SidSrc);
            var values = this._records.BuildValuesArray();

            if (stream.Array.Length == 0)
            {
                throw new Exceptions.InternalAssertionException();
            }

            if (type == Core.CellValueType.Formula)
            {
                var flags = this._compute_setformula_flags();
                var c = visobjtarget.SetFormulas(stream, values, (short) flags);
            }
            else
            {
                const object[] unitcodes = null;
                var flags = this._compute_setresults_flags();

                var res = visobjtarget.SetResults(stream, unitcodes, values, (short) flags);
            }
        }
    }
}