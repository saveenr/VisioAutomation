using VisioAutomation.ShapeSheet.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public class ShapeSheetWriter
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        private SIDSRCStreamBuilder FormulaRecords_SIDSRC;
        private ValuesBuilder FormulaRecords_SIDSRC_Formulas;

        private SRCStreamBuilder FormulaRecords_SRC;
        private ValuesBuilder FormulaRecords_SRC_Formulas;

        private SRCStreamBuilder ResultRecords_SRC;
        private ValuesBuilder ResultRecords_SRC_Results;
        private UnitCodesBuilder ResultRecords_SRC_UnitCodes;

        private SIDSRCStreamBuilder ResultRecords_SIDSRC;
        private ValuesBuilder ResultRecords_SIDSRC_Results;
        private UnitCodesBuilder ResultRecords_SIDSRC_UnitCodes;

        public ShapeSheetWriter()
        {
        }

        public void Clear()
        {
            if (this.FormulaRecords_SIDSRC != null) { this.FormulaRecords_SIDSRC.Clear(); }
            if (this.FormulaRecords_SIDSRC_Formulas != null) { this.FormulaRecords_SIDSRC_Formulas.Clear(); }

            if (this.FormulaRecords_SRC != null) { this.FormulaRecords_SRC.Clear(); }
            if (this.FormulaRecords_SRC_Formulas != null) { this.FormulaRecords_SRC_Formulas.Clear(); }

            if (this.ResultRecords_SRC != null) { this.ResultRecords_SRC.Clear(); }
            if (this.ResultRecords_SRC_Results != null) { this.ResultRecords_SRC_Results.Clear(); }
            if (this.ResultRecords_SRC_UnitCodes != null) { this.ResultRecords_SRC_UnitCodes.Clear(); }

            if (this.ResultRecords_SIDSRC != null) { this.ResultRecords_SIDSRC.Clear(); }
            if (this.ResultRecords_SIDSRC_Results != null) { this.ResultRecords_SIDSRC_Results.Clear(); }
            if (this.ResultRecords_SIDSRC_UnitCodes != null) { this.ResultRecords_SIDSRC_UnitCodes.Clear(); }

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

        public void SetFormula(SRC src, ValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(src, formula);
        }

        public void SetFormula(short id, SRC src, ValueLiteral formula)
        {
            var sidsrc = new SIDSRC(id, src);
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        public void SetFormula(SIDSRC sidsrc, ValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        private void __SetFormulaIgnoreNull(SRC src, ValueLiteral formula)
        {
            if (this.FormulaRecords_SRC == null)
            {
                this.FormulaRecords_SRC = new SRCStreamBuilder();
                this.FormulaRecords_SRC_Formulas = new ValuesBuilder();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_SRC.Add(src);
                this.FormulaRecords_SRC_Formulas.Add(formula.Value);
            }
        }

        private void __SetFormulaIgnoreNull(SIDSRC sidsrc, ValueLiteral formula)
        {
            if (this.FormulaRecords_SIDSRC == null)
            {
                this.FormulaRecords_SIDSRC = new SIDSRCStreamBuilder();
                this.FormulaRecords_SIDSRC_Formulas = new ValuesBuilder();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_SIDSRC.Add(sidsrc);
                this.FormulaRecords_SIDSRC_Formulas.Add(formula.Value);
            }
        }

        private void CommitFormulaRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            var stream_builder = coord_type == CoordType.SIDSRC ? (ShapeSheetStreamBuilder)this.FormulaRecords_SIDSRC : (ShapeSheetStreamBuilder)this.FormulaRecords_SRC;
            var formulas_builder = coord_type == CoordType.SIDSRC ? this.FormulaRecords_SIDSRC_Formulas : this.FormulaRecords_SRC_Formulas;

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

        public void SetResult(SRC src, ValueLiteral result, IVisio.VisUnitCodes unitcode)
        {
            if (this.ResultRecords_SRC == null)
            {
                this.ResultRecords_SRC = new SRCStreamBuilder();
                this.ResultRecords_SRC_Results = new ValuesBuilder();
                this.ResultRecords_SRC_UnitCodes = new UnitCodesBuilder();
            }

            this.ResultRecords_SRC.Add(src);
            this.ResultRecords_SRC_Results.Add(result.Value);
            this.ResultRecords_SRC_UnitCodes.Add(unitcode);
        }

        public void SetResult(short id, SRC src, ValueLiteral result, IVisio.VisUnitCodes unitcode)
        {
            var sidsrc = new SIDSRC(id, src);
            this.SetResult(sidsrc, result.Value, unitcode);
        }

        public void SetResult(SIDSRC sidsrc, ValueLiteral result, IVisio.VisUnitCodes unitcode)
        {
            if (this.ResultRecords_SIDSRC == null)
            {
                this.ResultRecords_SIDSRC = new SIDSRCStreamBuilder();
                this.ResultRecords_SIDSRC_Results = new ValuesBuilder();
                this.ResultRecords_SIDSRC_UnitCodes = new UnitCodesBuilder();
            }

            this.ResultRecords_SIDSRC.Add(sidsrc);
            this.ResultRecords_SIDSRC_Results.Add(result.Value);
            this.ResultRecords_SIDSRC_UnitCodes.Add(unitcode);

        }

        private void CommitResultRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            var stream_builder = coord_type == CoordType.SIDSRC ? (ShapeSheetStreamBuilder)this.ResultRecords_SIDSRC: (ShapeSheetStreamBuilder)this.ResultRecords_SRC;
            var results_builder = coord_type == CoordType.SIDSRC ? this.ResultRecords_SIDSRC_Results: this.ResultRecords_SRC_Results;
            var unitcodes_builder = coord_type == CoordType.SIDSRC ? this.ResultRecords_SIDSRC_UnitCodes: this.ResultRecords_SRC_UnitCodes;

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
            surface.SetResults(stream, unitcodes_builder.ToObjectArray(), results_builder.ToObjectArray(), (short)flags);
        }
    }
}