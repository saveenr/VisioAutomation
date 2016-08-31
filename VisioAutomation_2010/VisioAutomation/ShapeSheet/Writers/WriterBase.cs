using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public abstract class WriterBase<TStreamType, TValue>
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        public readonly List<TStreamType> StreamItems;
        public readonly List<TValue> ValueItems;

        public void Clear()
        {
            this.StreamItems.Clear();
            this.ValueItems.Clear();
        }

        protected WriterBase()
        {
            this.StreamItems = new List<TStreamType>();
            this.ValueItems = new List<TValue>();
        }

        protected WriterBase(int capacity)
        {
            this.StreamItems = new List<TStreamType>(capacity);
            this.ValueItems = new List<TValue>(capacity);
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

        protected abstract void _commit_to_surface(VisioAutomation.ShapeSheet.ShapeSheetSurface surface);

        public void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            this._commit_to_surface(surface);
        }
        public void Commit(IVisio.Shape shape)
        {
            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(shape);
            this._commit_to_surface(surface);                
        }

        public void Commit(IVisio.Page shape)
        {
            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(shape);
            this._commit_to_surface(surface);
        }

        public void Commit(IVisio.Master shape)
        {
            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(shape);
            this._commit_to_surface(surface);
        }

        public int Count
        {
            get { return this.ValueItems.Count; }
        }

    }
}
