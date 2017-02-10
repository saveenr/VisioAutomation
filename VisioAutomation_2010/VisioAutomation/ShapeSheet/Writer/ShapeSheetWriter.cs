using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class ShapeSheetWriter
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        private readonly List<WriteRecord<FormulaLiteral>> FormulaRecords;
        private readonly List<WriteRecord<ResultValue>> ResultRecords;

        public ShapeSheetWriter()
        {
            this.FormulaRecords = new List<WriteRecord<FormulaLiteral>>();
            this.ResultRecords = new List<WriteRecord<ResultValue>>();
        }

        public void Clear()
        {
            this.FormulaRecords.Clear();
            this.ResultRecords.Clear();
        }

        private void AddFormulaRecord(SRC src, FormulaLiteral value)
        {
            var rec = new WriteRecord<FormulaLiteral>(src, value);
            this.FormulaRecords.Add(rec);
        }

        private void AddFormulaRecord(SIDSRC sidsrc, FormulaLiteral value)
        {
            var rec = new WriteRecord<FormulaLiteral>(sidsrc, value);
            this.FormulaRecords.Add(rec);
        }

        private void AddResultRecord(SRC src, ResultValue value)
        {
            var rec = new WriteRecord<ResultValue>(src, value);
            this.ResultRecords.Add(rec);
        }

        private void AddResultRecord(SIDSRC sidsrc, ResultValue value)
        {
            var rec = new WriteRecord<ResultValue>(sidsrc, value);
            this.ResultRecords.Add(rec);
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
            this.CommitFormulaRecordsByType(surface, CoordType.SRC);
            this.CommitFormulaRecordsByType(surface, CoordType.SIDSRC);
            this.CommitResultRecordsByType(surface, CoordType.SRC);
            this.CommitResultRecordsByType(surface, CoordType.SIDSRC);
        }

        public int FormulaCount => this.FormulaRecords.Count;

        private IEnumerable<WriteRecord<FormulaLiteral>> GetFormulaRecords(CoordType type)
        {
            return this.FormulaRecords.Where(i => i.Type == type);
        }

        private IEnumerable<WriteRecord<ResultValue>> GetResultRecords(CoordType type)
        {
            return this.ResultRecords.Where(i => i.Type == type);
        }

        public void SetFormula(SRC src, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this.AddFormulaRecord(src,formula);
            }
        }

        public void SetFormula(short id, SRC src, FormulaLiteral formula)
        {
            var sidsrc = new SIDSRC(id, src);
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        public void SetFormula(SIDSRC sidsrc, FormulaLiteral formula)
        {
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        private void __SetFormulaIgnoreNull(SIDSRC sidsrc, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this.AddFormulaRecord(sidsrc,formula);
            }
        }

        private void CommitFormulaRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            var records = this.GetFormulaRecords(coord_type);
            var count = records.Count();

            if (count == 0)
            {
                return;
            }

            int chunksize = coord_type == CoordType.SIDSRC ? 4 : 3;

            var stream = new short[count * chunksize];
            var formulas = new object[count];

            int streampos = 0;
            int formulapos = 0;

            foreach (var rec in records)
            {
                // fill stream
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

                // fill formulas
                formulas[formulapos++] = rec.Value.Value;
            }

            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        public void SetResult(SRC src, string value, IVisio.VisUnitCodes unitcode)
        {
            var value_item = new ResultValue(value, unitcode);
            this.AddResultRecord(src, value_item);
        }

        public void SetResult(SRC src, double value, IVisio.VisUnitCodes unitcode)
        {
            var value_item = new ResultValue(value, unitcode);
            this.AddResultRecord(src, value_item);
        }

        public void SetResult(SIDSRC sidsrc, double value, IVisio.VisUnitCodes unitcode)
        {
            var v = new ResultValue(value, unitcode);
            this.AddResultRecord(sidsrc, v);
        }

        public void SetResult(SIDSRC sidsrc, string value, IVisio.VisUnitCodes unitcode)
        {
            var v = new ResultValue(value, unitcode);
            this.AddResultRecord(sidsrc, v);
        }

        private void CommitResultRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            var records = this.GetResultRecords(coord_type);
            var count = records.Count();

            if (count == 0)
            {
                return;
            }

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

                // fill results
                if (rec.Value.ResultType == ResultType.ResultNumeric)
                {
                    results[resultspos++] = rec.Value.ValueNumeric;
                }
                else if (rec.Value.ResultType == ResultType.ResultString)
                {
                    results[resultspos++] = rec.Value.ValueString;
                }

                // fill unit codes
                unitcodes[unitcodespos] = rec.Value.UnitCode;

            }

            var flags = this.ComputeGetResultFlags(records.First().Value.ResultType);
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }

}