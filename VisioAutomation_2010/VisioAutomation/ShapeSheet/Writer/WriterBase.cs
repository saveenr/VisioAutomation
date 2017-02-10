using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public abstract class WriterBase<TValue>
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        protected readonly List<WriteRecord<TValue>> Records;

        public void Clear()
        {
            this.Records.Clear();
        }

        protected void Add(SRC src, TValue value)
        {
            var rec = new WriteRecord<TValue>(src, value);
            this.Records.Add(rec);
        }

        protected void Add(SIDSRC sidsrc, TValue value)
        {
            var rec = new WriteRecord<TValue>(sidsrc, value);
            this.Records.Add(rec);
        }

        protected WriterBase()
        {
            this.Records = new List<WriteRecord<TValue>>();
        }

        protected IVisio.VisGetSetArgs ComputeGetResultFlags(ResultType rt)
        {
            var flags = this.combine_blastguards_and_testcircular_flags();

            if (rt == ResultType.ResultString)
            {
                flags |= IVisio.VisGetSetArgs.visGetStrings;
            }

            return flags;
        }

        protected IVisio.VisGetSetArgs ComputeGetFormulaFlags()
        {
            var common_flags = this.combine_blastguards_and_testcircular_flags();
            var formula_flags = (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;
            var combined_flags = (short)common_flags | formula_flags;
            return (IVisio.VisGetSetArgs)combined_flags;
        }

        private IVisio.VisGetSetArgs combine_blastguards_and_testcircular_flags()
        {
            var f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            var f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = ((short)f_bg) | ((short)f_tc);
            return (IVisio.VisGetSetArgs)flags;
        }

        public void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            this.CommitRecordsByType(surface, CoordType.SRC);
            this.CommitRecordsByType(surface, CoordType.SIDSRC);
        }

        protected abstract void CommitRecordsByType(VisioAutomation.ShapeSheet.ShapeSheetSurface surface, CoordType coord_type);

        public int Count => this.Records.Count;

        protected IEnumerable<WriteRecord<TValue>> GetRecords(CoordType type)
        {
            return this.Records.Where(i => i.Type == type);
        }
    }
}
