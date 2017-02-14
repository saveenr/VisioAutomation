using System.Linq;
using VisioAutomation.ShapeSheet.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public class ShapeSheetWriter
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        private readonly WriteRecords FormulaRecords;
        private readonly WriteRecords ResultRecords;

        public ShapeSheetWriter()
        {
            this.FormulaRecords = new WriteRecords();
            this.ResultRecords = new WriteRecords();
        }

        public void Clear()
        {
            this.FormulaRecords.Clear();
            this.ResultRecords.Clear();
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
            if (formula.HasValue)
            {
                this.FormulaRecords.Add(src, formula.Value, null);
            }
        }

        private void __SetFormulaIgnoreNull(SIDSRC sidsrc, ValueLiteral formula)
        {
            if (formula.HasValue)
            {
                this.FormulaRecords.Add(sidsrc, formula.Value, null);
            }
        }

        private void CommitFormulaRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            int count = this.FormulaRecords.CountByCoordType(coord_type);

            if (count == 0)
            {
                return;
            }

            var records = this.FormulaRecords.EnumerateByCoordType(coord_type);

            int chunksize = coord_type == CoordType.SIDSRC ? 4 : 3;

            var stream = new short[count * chunksize];
            var formulas = new object[count];

            int streampos = 0;
            int formulapos = 0;

            foreach (var rec in records)
            {
                // fill stream
                streampos = this.AddStreamRecord(stream, streampos, coord_type, rec);

                // fill formulas
                formulas[formulapos++] = rec.CellValue;

                if (rec.UnitCode.HasValue)
                {
                    throw new VisioAutomation.Exceptions.InternalAssertionException("Unit code should not be set for formulas");
                 }
            }

            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        public void SetResult(SRC src, ValueLiteral result, IVisio.VisUnitCodes unitcode)
        {
            this.ResultRecords.Add(src, result.Value, unitcode);
        }

        public void SetResult(short id, SRC src, ValueLiteral result, IVisio.VisUnitCodes unitcode)
        {
            var sidsrc = new SIDSRC(id, src);
            this.ResultRecords.Add(sidsrc, result.Value, unitcode);
        }

        public void SetResult(SIDSRC sidsrc, ValueLiteral result, IVisio.VisUnitCodes unitcode)
        {
            this.ResultRecords.Add(sidsrc, result.Value, unitcode);
        }

        private void CommitResultRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {

            int count = this.ResultRecords.CountByCoordType(coord_type);

            if (count == 0)
            {
                return;
            }

            var records = this.ResultRecords.EnumerateByCoordType(coord_type);

            int chunksize = coord_type == CoordType.SIDSRC ? 4 : 3;

            var stream = new short[count * chunksize];
            var results = new object[count];
            var unitcodes = new object[count];

            int streampos = 0;
            int resultspos = 0;
            int unitcodespos = 0;

            foreach (var rec in records)
            {
                // fill stream
                streampos = this.AddStreamRecord(stream, streampos, coord_type, rec);

                // fill results
                results[resultspos++] = rec.CellValue;

                // fill unit codes
                if (!rec.UnitCode.HasValue)
                {
                    throw new VisioAutomation.Exceptions.InternalAssertionException("Unit code missing for result");
                }
                unitcodes[unitcodespos++] = rec.UnitCode.Value;
            }

            var flags = this.ComputeGetResultFlags();
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }

        private int AddStreamRecord(short[] stream, int streampos, CoordType coord_type, WriteRecord rec)
        {
            if (coord_type == CoordType.SRC)
            {
                var src = rec.SRC;
                stream[streampos++] = src.Section;
                stream[streampos++] = src.Row;
                stream[streampos++] = src.Cell;
            }
            else
            {
                var sidsrc = rec.SIDSRC;
                stream[streampos++] = sidsrc.ShapeID;
                stream[streampos++] = sidsrc.Section;
                stream[streampos++] = sidsrc.Row;
                stream[streampos++] = sidsrc.Cell;
            }
            return streampos;
        }
    }

}