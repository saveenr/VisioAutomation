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
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        private UpdateRecord? first_record;

        public void Clear()
        {
            this.items.Clear();
            this.first_record = null;
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
            // This block ensures that only homogeneous updates are constructed
            if (this.items.Count<1)
            {
                this.first_record = record;
            }
            else
            {
                if (record.StreamType != first_record.Value.StreamType)
                {
                    throw new VA.AutomationException("Cannot contain both SRC and SIDSRC updates");
                }

                if (record.UpdateType != first_record.Value.UpdateType)
                {
                    throw new VA.AutomationException("Cannot contain both Formula and Result updates");
                }
            }

            // Now that it is safe, add the record
            this.items.Add(record);

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
            var a = new string[this.items.Count];
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
            var a = new double[this.items.Count];
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
            var a = new IVisio.VisUnitCodes[this.items.Count];
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
            this._Execute(page);
        }

        public void Execute(IVisio.Shape shape)
        {
            this._Execute(shape);
        }

        private void _Execute(object visio_object)
        {
            if (!(visio_object is IVisio.Page || visio_object is IVisio.Shape))
            {
                throw new VA.AutomationException();
            }

            if (this.items.Count < 1)
            {
                return;
            }

            if (visio_object is IVisio.Shape)
            {
                if (first_record.Value.StreamType == StreamType.SIDSRC)
                {
                    throw new VA.AutomationException("Contains a SIDSRC updates. Need SRC updates");
                }
            }
            else if (visio_object is IVisio.Page)
            {
                if (first_record.Value.StreamType == StreamType.SRC)
                {
                    throw new VA.AutomationException("Contains a SRC updates. Need SIDSRC updates");
                }
            }
            
            if (first_record.Value.UpdateType == UpdateType.Result)
            {

                var stream = this.build_stream();
                var unitcodes = this.GetUnitCodesArray();
                var results = this.GetResultsArray();
                var flags = this.ResultFlags;

                if (visio_object is IVisio.Shape)
                {
                    int c = VA.ShapeSheet.ShapeSheetHelper.SetResults( (IVisio.Shape) visio_object, stream, results, unitcodes, flags);                    
                }
                else if (visio_object is IVisio.Page)
                {
                    int c = VA.ShapeSheet.ShapeSheetHelper.SetResults( (IVisio.Page) visio_object, stream, results, unitcodes, flags);
                }
            }
            else
            {
                var stream = this.build_stream();
                var formulas = this.GetFormulasArray();
                var flags = this.FormulaFlags;
                if (visio_object is IVisio.Shape)
                {
                    int c = VA.ShapeSheet.ShapeSheetHelper.SetFormulas((IVisio.Shape) visio_object, stream, formulas, flags);
                }
                else if (visio_object is IVisio.Page)
                {
                    int c = VA.ShapeSheet.ShapeSheetHelper.SetFormulas( (IVisio.Page) visio_object, stream, formulas, (short) flags);
                }

            }
        }


        private short [] build_stream()
        {
            var st = this.first_record.Value.StreamType;

            if (st==StreamType.SRC)
            {
                var streamb = new List<SRC>(this.items.Count);
                streamb.AddRange( this.items.Select(i=>i.SIDSRC.SRC));
                return SRC.ToStream(streamb);
            }
            else
            {
                var streamb = new List<SIDSRC>(this.items.Count);
                streamb.AddRange(this.items.Select(i => i.SIDSRC));
                return SIDSRC.ToStream(streamb);
            }
            
        }

        public void SetFormula(SRC streamitem, FormulaLiteral formula)
        {
            this._SetFormula(StreamType.SRC, new SIDSRC(-1, streamitem), formula);
        }

        public void SetFormulaIgnoreNull(SRC streamitem, ShapeSheet.FormulaLiteral formula)
        {
            this._SetFormulaIgnoreNull(StreamType.SRC, new SIDSRC(-1, streamitem), formula);
        }

        public void SetResult(SRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(StreamType.SRC, new SIDSRC(-1, streamitem), value, unitcode);
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