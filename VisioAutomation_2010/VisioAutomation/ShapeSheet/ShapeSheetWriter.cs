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

        public void SetFormula(SRC src, CellValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(src, formula);
        }

        public void SetFormula(short id, SRC src, CellValueLiteral formula)
        {
            var sidsrc = new SIDSRC(id, src);
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        public void SetFormula(SIDSRC sidsrc, CellValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        private void __SetFormulaIgnoreNull(SRC src, CellValueLiteral formula)
        {
            if (this.FormulaRecords_SRC == null)
            {
                this.FormulaRecords_SRC = new WriterCollection_SRC(false);
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_SRC.StreamBuilder.Add(src);
                this.FormulaRecords_SRC.ValuesBuilder.Add(formula.Value);
            }
        }

        private void __SetFormulaIgnoreNull(SIDSRC sidsrc, CellValueLiteral formula)
        {
            if (this.FormulaRecords_SIDSRC == null)
            {
                this.FormulaRecords_SIDSRC = new WriterCollection_SIDSRC(false);
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_SIDSRC.StreamBuilder.Add(sidsrc);
                this.FormulaRecords_SIDSRC.ValuesBuilder.Add(formula.Value);
            }
        }

        private void CommitFormulaRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            if (coord_type == CoordType.SIDSRC && this.FormulaRecords_SIDSRC == null)
            {
                return;
            }

            if (coord_type == CoordType.SRC && this.FormulaRecords_SRC == null)
            {
                return;
            }


            var stream_builder = coord_type == CoordType.SIDSRC ? (ShapeSheetStreamBuilder) this.FormulaRecords_SIDSRC.StreamBuilder : (ShapeSheetStreamBuilder)this.FormulaRecords_SRC.StreamBuilder;
            var formulas_builder = coord_type == CoordType.SIDSRC ? this.FormulaRecords_SIDSRC.ValuesBuilder : this.FormulaRecords_SRC.ValuesBuilder;

            if (formulas_builder == null)
            {
                return;
            }

            int count = formulas_builder.Count;

            if (count == 0)
            {
                return;
            }

            var stream = stream_builder.ToStream();

            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas_builder.ToObjectArray(), (short)flags);
        }

        public void SetResult(SRC src, CellValueLiteral result, IVisio.VisUnitCodes unitcode)
        {
            if (this.ResultRecords_SRC == null)
            {
                this.ResultRecords_SRC = new WriterCollection_SRC(true);
            }

            this.ResultRecords_SRC.StreamBuilder.Add(src);
            this.ResultRecords_SRC.ValuesBuilder.Add(result.Value);
            this.ResultRecords_SRC.UnitCodesBuilder.Add(unitcode);
        }

        public void SetResult(short id, SRC src, CellValueLiteral result, IVisio.VisUnitCodes unitcode)
        {
            var sidsrc = new SIDSRC(id, src);
            this.SetResult(sidsrc, result.Value, unitcode);
        }

        public void SetResult(SIDSRC sidsrc, CellValueLiteral result, IVisio.VisUnitCodes unitcode)
        {
            if (this.ResultRecords_SIDSRC == null)
            {
                this.ResultRecords_SIDSRC = new WriterCollection_SIDSRC(true);
            }

            this.ResultRecords_SIDSRC.StreamBuilder.Add(sidsrc);
            this.ResultRecords_SIDSRC.ValuesBuilder.Add(result.Value);
            this.ResultRecords_SIDSRC.UnitCodesBuilder.Add(unitcode);
        }

        private void CommitResultRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            if (coord_type == CoordType.SIDSRC && this.ResultRecords_SIDSRC == null)
            {
                return;
            }

            if (coord_type == CoordType.SRC && this.ResultRecords_SRC == null)
            {
                return;
            }


            var stream_builder = coord_type == CoordType.SIDSRC ? (ShapeSheetStreamBuilder)this.ResultRecords_SIDSRC.StreamBuilder: (ShapeSheetStreamBuilder)this.ResultRecords_SRC.StreamBuilder;
            var results_builder = coord_type == CoordType.SIDSRC ? this.ResultRecords_SIDSRC.ValuesBuilder: this.ResultRecords_SRC.ValuesBuilder;
            var unitcodes_builder = coord_type == CoordType.SIDSRC ? this.ResultRecords_SIDSRC.UnitCodesBuilder: this.ResultRecords_SRC.UnitCodesBuilder;

            if (results_builder == null)
            {
                return;
            }

            int count = results_builder.Count;

            if (count == 0)
            {
                return;
            }

            var stream = stream_builder.ToStream();
            var flags = this.ComputeGetResultFlags();
            var unitcodes = unitcodes_builder.ToObjectArray();
            var results = results_builder.ToObjectArray();
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }
}