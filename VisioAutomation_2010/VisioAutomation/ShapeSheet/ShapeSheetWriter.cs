using VisioAutomation.ShapeSheet.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public class ShapeSheetWriter
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        private WriterCollection_SIDSRC FormulaRecords_SIDSRC;
        private WriterCollection_SRC FormulaRecords_SRC;
        private WriterCollection_SRC ResultRecords_SRC;
        private WriterCollection_SIDSRC ResultRecords_SIDSRC;

        public ShapeSheetWriter()
        {
        }

        public void Clear()
        {
            FormulaRecords_SIDSRC?.Clear();
            FormulaRecords_SRC?.Clear();
            ResultRecords_SRC?.Clear();
            ResultRecords_SIDSRC?.Clear();
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

        public void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            this.CommitFormulaRecordsByType(surface, CoordType.SRC);
            this.CommitFormulaRecordsByType(surface, CoordType.SIDSRC);
            this.CommitResultRecordsByType(surface, CoordType.SRC);
            this.CommitResultRecordsByType(surface, CoordType.SIDSRC);
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
            if (this.FormulaRecords_SRC == null)
            {
                this.FormulaRecords_SRC = new WriterCollection_SRC();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_SRC.StreamBuilder.Add(src);
                this.FormulaRecords_SRC.ValuesBuilder.Add(formula.Value);
            }
        }

        private void __SetFormulaIgnoreNull(SidSrc sidsrc, CellValueLiteral formula)
        {
            if (this.FormulaRecords_SIDSRC == null)
            {
                this.FormulaRecords_SIDSRC = new WriterCollection_SIDSRC();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_SIDSRC.Add(sidsrc, formula.Value);
            }
        }

        private void CommitFormulaRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            if (coord_type == CoordType.SIDSRC && (this.FormulaRecords_SIDSRC == null || this.FormulaRecords_SIDSRC.Count <1))
            {
                return;
            }

            if (coord_type == CoordType.SRC && (this.FormulaRecords_SRC == null || this.FormulaRecords_SRC.Count <1))
            {
                return;
            }

            var stream = coord_type == CoordType.SIDSRC ? this.FormulaRecords_SIDSRC.BuildStream() : this.FormulaRecords_SRC.BuildStream();
            var formulas = coord_type == CoordType.SIDSRC ? this.FormulaRecords_SIDSRC.BuildValues() : this.FormulaRecords_SRC.BuildValues();

            if (stream.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetFormulaFlags();

            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        public void SetResult(Src src, CellValueLiteral result)
        {
            if (this.ResultRecords_SRC == null)
            {
                this.ResultRecords_SRC = new WriterCollection_SRC();
            }

            this.ResultRecords_SRC.StreamBuilder.Add(src);
            this.ResultRecords_SRC.ValuesBuilder.Add(result.Value);
        }

        public void SetResult(short id, Src src, CellValueLiteral result)
        {
            var sidsrc = new SidSrc(id, src);
            this.SetResult(sidsrc, result.Value);
        }

        public void SetResult(SidSrc sidsrc, CellValueLiteral result)
        {
            if (this.ResultRecords_SIDSRC == null)
            {
                this.ResultRecords_SIDSRC = new WriterCollection_SIDSRC();
            }

            this.ResultRecords_SIDSRC.Add(sidsrc, result.Value);
        }

        private void CommitResultRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            if (coord_type == CoordType.SIDSRC && (this.ResultRecords_SIDSRC == null || this.ResultRecords_SIDSRC.Count < 1))
            {
                return;
            }

            if (coord_type == CoordType.SRC && (this.ResultRecords_SRC == null || this.ResultRecords_SRC.Count <1))
            {
                return;
            }

            var stream = coord_type == CoordType.SIDSRC ? this.ResultRecords_SIDSRC.BuildStream() : this.ResultRecords_SRC.BuildStream();
            var results = coord_type == CoordType.SIDSRC ? this.ResultRecords_SIDSRC.BuildValues(): this.ResultRecords_SRC.BuildValues();
            const object[] unitcodes = null;
            
            if (stream.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetResultFlags();
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }
}