using VisioAutomation.ShapeSheet.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public class ShapeSheetWriter
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        private WriterCollection_SidSrc FormulaRecords_SidSrc;
        private WriterCollection_Src FormulaRecords_Src;
        private WriterCollection_Src ResultRecords_Src;
        private WriterCollection_SidSrc ResultRecords_SidSrc;

        public ShapeSheetWriter()
        {
        }

        public void Clear()
        {
            FormulaRecords_SidSrc?.Clear();
            FormulaRecords_Src?.Clear();
            ResultRecords_Src?.Clear();
            ResultRecords_SidSrc?.Clear();
        }

        protected IVisio.VisGetSetArgs ComputeGetResultFlags()
        {
            var flags = this.combine_blastguards_and_testcircular_flags();

            flags |= IVisio.VisGetSetArgs.visGetStrings;

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

        public void Commit(IVisio.Shape shape)
        {
            var surface = new ShapeSheetSurface(shape);
            this.Commit(surface);
        }

        public void Commit(IVisio.Page page)
        {
            var surface = new ShapeSheetSurface(page);
            this.Commit(surface);
        }

        public void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            this.CommitFormulaRecordsByType(surface, CoordType.Src);
            this.CommitFormulaRecordsByType(surface, CoordType.SidSrc);
            this.CommitResultRecordsByType(surface, CoordType.Src);
            this.CommitResultRecordsByType(surface, CoordType.SidSrc);
        }

        public void SetFormula(Src src, CellValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(src, formula);
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

        private void __SetFormulaIgnoreNull(Src src, CellValueLiteral formula)
        {
            if (this.FormulaRecords_Src == null)
            {
                this.FormulaRecords_Src = new WriterCollection_Src();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_Src.StreamBuilder.Add(src);
                this.FormulaRecords_Src.ValuesBuilder.Add(formula.Value);
            }
        }

        private void __SetFormulaIgnoreNull(SidSrc sidsrc, CellValueLiteral formula)
        {
            if (this.FormulaRecords_SidSrc == null)
            {
                this.FormulaRecords_SidSrc = new WriterCollection_SidSrc();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_SidSrc.Add(sidsrc, formula.Value);
            }
        }

        private void CommitFormulaRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            if (coord_type == CoordType.SidSrc && (this.FormulaRecords_SidSrc == null || this.FormulaRecords_SidSrc.Count <1))
            {
                return;
            }

            if (coord_type == CoordType.Src && (this.FormulaRecords_Src == null || this.FormulaRecords_Src.Count <1))
            {
                return;
            }

            var stream = coord_type == CoordType.SidSrc ? this.FormulaRecords_SidSrc.BuildStream() : this.FormulaRecords_Src.BuildStream();
            var formulas = coord_type == CoordType.SidSrc ? this.FormulaRecords_SidSrc.BuildValues() : this.FormulaRecords_Src.BuildValues();

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetFormulaFlags();

            int c = surface.SetFormulas(stream.Array, formulas, (short)flags);
        }

        public void SetResult(Src src, CellValueLiteral result)
        {
            if (this.ResultRecords_Src == null)
            {
                this.ResultRecords_Src = new WriterCollection_Src();
            }

            this.ResultRecords_Src.StreamBuilder.Add(src);
            this.ResultRecords_Src.ValuesBuilder.Add(result.Value);
        }

        public void SetResult(short id, Src src, CellValueLiteral result)
        {
            var sidsrc = new SidSrc(id, src);
            this.SetResult(sidsrc, result.Value);
        }

        public void SetResult(SidSrc sidsrc, CellValueLiteral result)
        {
            if (this.ResultRecords_SidSrc == null)
            {
                this.ResultRecords_SidSrc = new WriterCollection_SidSrc();
            }

            this.ResultRecords_SidSrc.Add(sidsrc, result.Value);
        }

        private void CommitResultRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            if (coord_type == CoordType.SidSrc && (this.ResultRecords_SidSrc == null || this.ResultRecords_SidSrc.Count < 1))
            {
                return;
            }

            if (coord_type == CoordType.Src && (this.ResultRecords_Src == null || this.ResultRecords_Src.Count <1))
            {
                return;
            }

            var stream = coord_type == CoordType.SidSrc ? this.ResultRecords_SidSrc.BuildStream() : this.ResultRecords_Src.BuildStream();
            var results = coord_type == CoordType.SidSrc ? this.ResultRecords_SidSrc.BuildValues(): this.ResultRecords_Src.BuildValues();
            const object[] unitcodes = null;
            
            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetResultFlags();
            surface.SetResults(stream.Array, unitcodes, results, (short)flags);
        }
    }
}