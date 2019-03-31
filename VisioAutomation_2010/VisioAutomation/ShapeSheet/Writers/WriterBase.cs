using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class WriterBase
    {
        protected WriteRecordList _records;
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        protected WriterBase(StreamType type)
        {
            this._records = new WriteRecordList(type);
        }

        protected IVisio.VisGetSetArgs ComputeGetResultFlags()
        {
            var flags = this._combine_blastguards_and_testcircular_flags();

            flags |= IVisio.VisGetSetArgs.visGetStrings;

            return flags;
        }

        public void Clear()
        {
            _records.Clear();
        }


        protected IVisio.VisGetSetArgs _compute_get_formula_flags()
        {
            var common_flags = this._combine_blastguards_and_testcircular_flags();
            var formula_flags = (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;
            var combined_flags = (short)common_flags | formula_flags;
            return (IVisio.VisGetSetArgs)combined_flags;
        }

        private IVisio.VisGetSetArgs _combine_blastguards_and_testcircular_flags()
        {
            var f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            var f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = ((short)f_bg) | ((short)f_tc);
            return (IVisio.VisGetSetArgs)flags;
        }

    }
}