using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections;
using System.Linq;

namespace VisioAutomation.ShapeSheet
{
    public class Update : IEnumerable<Update.UpdateRecord>
    {
        private List<UpdateRecord> items;
        public int ResultCount { get; private set; }
        public int FormulaCount { get; private set; }
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }
        private bool contains_SIDSRC;
        private bool contains_SID;

        public void Clear()
        {
            this.items.Clear();
            this.FormulaCount = 0;
            this.ResultCount = 0;
        }

        public Update()
        {
            this.items = new List<UpdateRecord>();
        }

        public Update(int capacity)
        {
            this.items = new List<UpdateRecord>(capacity);
        }

        protected IVisio.VisGetSetArgs ResultFlags
        {
            get { return get_common_flags(); }
        }

        protected IVisio.VisGetSetArgs FormulaFlags
        {
            get
            {
                var common_flags = get_common_flags();
                var formula_flags = (short) IVisio.VisGetSetArgs.visSetUniversalSyntax;
                var combined_flags = (short) common_flags | formula_flags;
                return (IVisio.VisGetSetArgs) combined_flags;
            }
        }

        private IVisio.VisGetSetArgs get_common_flags()
        {
            IVisio.VisGetSetArgs f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            IVisio.VisGetSetArgs f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = (short) f_bg | (short) f_tc;
            return (IVisio.VisGetSetArgs) flags;
        }


        private void CheckFormulaIsNotNull(string formula)
        {
            if (formula == null)
            {
                throw new AutomationException("Null not allowed for formula");
            }
        }

        private void AddRecord(UpdateRecord record)
        {
            if (this.contains_SID && record.StreamType==StreamType.SIDSRC)
            {
                throw new VA.AutomationException("Cannot mix SIDSRC and SRC Update records");
            }
            else if (this.contains_SIDSRC && record.StreamType==StreamType.SRC)
            {
                throw new VA.AutomationException("Cannot mix SIDSRC and SRC Update records");                
            }

            if (record.StreamType==StreamType.SIDSRC)
            {
                this.contains_SIDSRC = true;
            }
            else
            {
                this.contains_SID = true;
            }

            this.items.Add(record);

            if (record.UpdateType == UpdateType.Result)
            {
                this.ResultCount++;
            }
            else
            {
                this.FormulaCount++;
            }
        }
        protected void _SetFormula(StreamType st,SIDSRC streamitem, FormulaLiteral formula)
        {
            this.CheckFormulaIsNotNull(formula.Value);
            var rec = new UpdateRecord(st, streamitem, formula.Value);
            this.AddRecord(rec);
        }

        protected void _SetFormulaIgnoreNull(StreamType st, SIDSRC streamitem, ShapeSheet.FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this._SetFormula(st, streamitem, formula);
            }
        }

        protected void _SetResult(StreamType st, SIDSRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            var rec = new UpdateRecord(st,streamitem, value, unitcode);
            this.AddRecord(rec);
        }

        public IEnumerator<UpdateRecord> GetEnumerator()
        {
            foreach (var i in this.items)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator() // Explicit implementation
        {
            // keeps it hidden.
            return GetEnumerator();
        }

        public IEnumerable<UpdateRecord> ResultRecords
        {
            get { return this.items.Where(i => i.UpdateType == UpdateType.Result); }
        }

        public IEnumerable<UpdateRecord> FormulaRecords
        {
            get { return this.items.Where(i => i.UpdateType == UpdateType.Formula); }
        }

        protected string[] GetFormulasArray()
        {
            var a = new string[this.FormulaCount];
            int i = 0;
            foreach (var rec in this.FormulaRecords)
            {
                a[i] = rec.Formula;
                i++;
            }
            return a;
        }

        protected double[] GetResultsArray()
        {
            var a = new double[this.ResultCount];
            int i = 0;
            foreach (var rec in this.ResultRecords)
            {
                a[i] = rec.Result;
                i++;
            }
            return a;
        }

        protected IVisio.VisUnitCodes[] GetUnitCodesArray()
        {
            var a = new IVisio.VisUnitCodes[this.ResultCount];
            int i = 0;
            foreach (var rec in this.ResultRecords)
            {
                a[i] = rec.UnitCode;
                i++;
            }
            return a;
        }

        public void SetResult(short shapeid, SRC src, double value, IVisio.VisUnitCodes unitcode)
        {
            var streamitem = new SIDSRC(shapeid,src);
            this._SetResult(StreamType.SIDSRC, streamitem, value, unitcode);
        }

        public void SetResult(SIDSRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(StreamType.SIDSRC, streamitem, value, unitcode);
        }

        public void SetFormula(SIDSRC streamitem, FormulaLiteral formula)
        {
            this._SetFormula(StreamType.SIDSRC, streamitem, formula);
        }

        public void SetFormula(short shapeid, SRC src, FormulaLiteral formula)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this._SetFormula(StreamType.SIDSRC, streamitem, formula);
        }

        public void SetFormulaIgnoreNull(short id, ShapeSheet.SRC src, ShapeSheet.FormulaLiteral formula)
        {
            var sidsrc = new VA.ShapeSheet.SIDSRC(id, src);
            this._SetFormulaIgnoreNull(StreamType.SIDSRC, sidsrc, formula);
        }

        public void Execute(IVisio.Page page)
        {
            this.SetResults(page);
            this.SetFormulas(page);
        }

        private short SetResults( IVisio.Page page)
        {
            if (this.ResultCount== 0)
            {
                return 0;
            }

            var streamb = new List<SIDSRC>(this.ResultCount);
            streamb.AddRange(this.ResultRecords.Select(i => i.SIDSRC));
            var stream = SIDSRC.ToStream(streamb);

            var unitcodes = this.GetUnitCodesArray();
            double[] results = this.GetResultsArray();
            var flags = this.ResultFlags;

            return VA.ShapeSheet.ShapeSheetHelper.SetResults(page, stream, results, unitcodes, flags);
        }


        private short SetFormulas(IVisio.Page page)
        {
            if (this.FormulaCount == 0)
            {
                return 0;
            }

            var streamb = new List<SIDSRC>(this.FormulaCount);
            streamb.AddRange(this.FormulaRecords.Select(i => i.SIDSRC));
            var stream = SIDSRC.ToStream(streamb); 
            var formulas = this.GetFormulasArray();
            var flags = this.FormulaFlags;

            return VA.ShapeSheet.ShapeSheetHelper.SetFormulas(page, stream, formulas, (short)flags);
        }

        public void Execute(IVisio.Shape shape)
        {
            this.SetResults(shape);
            this.SetFormulas(shape);
        }

        private short SetResults(
            IVisio.Shape shape)
        {
            if (this.ResultCount== 0)
            {
                return 0;
            }

            var streamb = new List<SRC>(this.ResultCount);
            streamb.AddRange(this.ResultRecords.Select(i => i.SIDSRC.SRC));
            var stream = SRC.ToStream(streamb); 
            var unitcodes = this.GetUnitCodesArray();
            var results = this.GetResultsArray();
            var flags = this.ResultFlags;
            return VA.ShapeSheet.ShapeSheetHelper.SetResults(shape, stream, results, unitcodes, flags);
        }

        public void SetFormula(SRC streamitem, FormulaLiteral formula)
        {
            this._SetFormula(StreamType.SIDSRC, new SIDSRC(-1, streamitem), formula);
        }

        public void SetFormulaIgnoreNull(SRC streamitem, ShapeSheet.FormulaLiteral formula)
        {
            this._SetFormulaIgnoreNull(StreamType.SIDSRC, new SIDSRC(-1, streamitem), formula);
        }

        public void SetResult(SRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(StreamType.SIDSRC, new SIDSRC(-1, streamitem), value, unitcode);
        }

        private short SetFormulas(IVisio.Shape shape)
        {
            if (this.FormulaCount == 0)
            {
                return 0;
            }

            var streamb = new List<SRC>(this.FormulaCount);
            streamb.AddRange(this.FormulaRecords.Select(i => i.SIDSRC.SRC));
            var stream = SRC.ToStream(streamb); 
            var formulas = this.GetFormulasArray();
            var flags = this.FormulaFlags;
            return VA.ShapeSheet.ShapeSheetHelper.SetFormulas(shape, stream, formulas, flags);
        }

        public struct UpdateRecord
        {
            public readonly SIDSRC SIDSRC;
            public readonly string Formula;
            public readonly double Result;
            public readonly IVisio.VisUnitCodes UnitCode;
            public readonly UpdateType UpdateType;
            public readonly StreamType StreamType;

            internal UpdateRecord(StreamType st, SIDSRC sidsrc, string formula)
            {
                this.SIDSRC = sidsrc;
                this.Formula = formula;
                this.Result = 0.0;
                this.UnitCode = IVisio.VisUnitCodes.visNoCast;
                this.UpdateType = UpdateType.Formula;
                this.StreamType = st;
            }

            internal UpdateRecord(StreamType st, SIDSRC sidsrc, double result, IVisio.VisUnitCodes unit_code)
            {
                this.SIDSRC = sidsrc;
                this.Formula = null;
                this.UnitCode = unit_code;
                this.Result = result;
                this.UpdateType = UpdateType.Result;
                this.StreamType = st;
            }

        }

        public enum StreamType
        {
            SIDSRC, SRC
        }

        public enum UpdateType
        {
            Formula,
            Result
        }
    }
}