using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections;
using System.Linq;

namespace VisioAutomation.ShapeSheet
{
    public class Update : IEnumerable<Update.UpdateRecord>
    {
        private readonly List<UpdateRecord> updates;
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        private UpdateRecord? first_update;

        public void Clear()
        {
            this.updates.Clear();
            this.first_update = null;
        }

        public Update()
        {
            this.updates = new List<UpdateRecord>();
        }

        public Update(int capacity)
        {
            this.updates = new List<UpdateRecord>(capacity);
        }

        protected IVisio.VisGetSetArgs ResultFlags
        {
            get 
            { 
                var flags = get_common_flags();
                if ((flags & IVisio.VisGetSetArgs.visSetFormulas) > 0)
                {
                    flags = (IVisio.VisGetSetArgs)((short)flags | (short)IVisio.VisGetSetArgs.visSetUniversalSyntax);
                }
                return flags;
            }
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

        private void _add_update(UpdateRecord update)
        {
            // This block ensures that only homogeneous updates are constructed
            if (!this.first_update.HasValue)
            {
                this.first_update = update;
            }
            else
            {
                // first validate the stream types
                if (first_update.Value.StreamType != update.StreamType)
                {
                    throw new VA.AutomationException("Cannot contain both SRC and SIDSRC updates");
                }

                // Now ensure that we aren't mixing formulas and results
                // Keep in mind that we can mix differnt types of results (strings and numerics)
                if (first_update.Value.UpdateType == UpdateType.Formula && update.UpdateType != UpdateType.Formula)
                {
                    if (update.UpdateType != UpdateType.Formula)
                    {
                        throw new VA.AutomationException("Cannot contain both Formula and Result updates");
                    }
                }
                else if (first_update.Value.UpdateType == UpdateType.ResultNumeric || first_update.Value.UpdateType == UpdateType.ResultString)
                {
                    if (update.UpdateType == UpdateType.Formula)
                    {
                        throw new VA.AutomationException("Cannot contain both Formula and Result updates");
                    }
                }
            }

            // Now that it is safe, add the record
            this.updates.Add(update);

        }
        protected void _SetFormula(StreamType st,SIDSRC streamitem, FormulaLiteral formula)
        {
            this.CheckFormulaIsNotNull(formula.Value);
            var rec = new UpdateRecord(st, streamitem, formula.Value);
            this._add_update(rec);
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
            this._add_update(rec);
        }

        protected void _SetResult(StreamType st, SIDSRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            var rec = new UpdateRecord(st, streamitem, value, unitcode);
            this._add_update(rec);
        }

        public IEnumerator<UpdateRecord> GetEnumerator()
        {
            foreach (var i in this.updates)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator() 
        {
            // keeps it hidden.
            return GetEnumerator();
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

        public void SetResult(short shapeid, SRC src, string value, IVisio.VisUnitCodes unitcode)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this._SetResult(StreamType.SIDSRC, streamitem, value, unitcode);
        }

        public void SetResult(SIDSRC streamitem, string value, IVisio.VisUnitCodes unitcode)
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

        public void SetFormulas(VA.ShapeSheet.CellGroups.CellGroup cg)
        {
            foreach (var pair in cg.Pairs())
            {
                this.SetFormulaIgnoreNull(pair.SRC, pair.Formula);
            }
        }

        public void SetFormulas(short shapeid, VA.ShapeSheet.CellGroups.CellGroup cg)
        {
            foreach (var pair in cg.Pairs())
            {
                this.SetFormulaIgnoreNull(shapeid, pair.SRC, pair.Formula);
            }
        }

        public void SetFormulas(short shapeid, VA.ShapeSheet.CellGroups.CellGroupMultiRow cg, short row)
        {
            foreach (var pair in cg.Pairs())
            {
                this.SetFormulaIgnoreNull(shapeid, pair.SRC.ForRow(row), pair.Formula);            
            }
        }

        public void SetFormulas(VA.ShapeSheet.CellGroups.CellGroupMultiRow cg, short row)
        {
            foreach (var pair in cg.Pairs())
            {
                this.SetFormulaIgnoreNull(pair.SRC.ForRow(row), pair.Formula);
            }
        }
        
        public void Execute(IVisio.Page page)
        {
            var surface = new ShapeSheetSurface(page);
            this._Execute(surface);
        }

        public void Execute(IVisio.Shape shape)
        {
            var surface = new ShapeSheetSurface(shape);
            this._Execute(surface);
        }

        public void Execute(ShapeSheetSurface surface)
        {
            this._Execute(surface);
        }

        private void _Execute(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.updates.Count < 1)
            {
                return;
            }

            if (surface.Target.Shape != null)
            {
                if (first_update.Value.StreamType == StreamType.SIDSRC)
                {
                    throw new VA.AutomationException("Contains a SIDSRC updates. Need SRC updates");
                }
            }
            else if (surface.Target.Master != null)
            {
                if (first_update.Value.StreamType == StreamType.SIDSRC)
                {
                    throw new VA.AutomationException("Contains a SIDSRC updates. Need SRC updates");
                }
            }
            else if (surface.Target.Page != null)
            {
                if (first_update.Value.StreamType == StreamType.SRC)
                {
                    throw new VA.AutomationException("Contains a SRC updates. Need SIDSRC updates");
                }
            }

            var stream = this.build_stream();

            if (first_update.Value.UpdateType == UpdateType.ResultNumeric || first_update.Value.UpdateType==UpdateType.ResultString)
            {
                // Set Results

                // Create the unitcodes and results arrays
                var unitcodes = new object[this.updates.Count];
                var results = new object[this.updates.Count];
                int i = 0;
                foreach (var update in this.updates)
                {
                    unitcodes[i] = update.UnitCode;
                    if (update.UpdateType == UpdateType.ResultNumeric)
                    {
                        results[i] = update.ResultNumeric;                       
                    }
                    else if (update.UpdateType == UpdateType.ResultString)
                    {
                        results[i] = update.ResultString;
                    }
                    else
                    {
                        throw new AutomationException("Unhandled update type");
                    }
                    i++;
                }
                
                var flags = this.ResultFlags;

                if (first_update.Value.UpdateType == UpdateType.ResultNumeric)
                {
                }
                else if (first_update.Value.UpdateType == UpdateType.ResultString)
                {
                    flags |= IVisio.VisGetSetArgs.visGetStrings;
                }

                surface.SetResults(stream, unitcodes, results, (short)flags);                    
            }
            else
            {
                // Set Formulas

                // Create the formulas array
                var formulas = new object[this.updates.Count];
                int i = 0;
                foreach (var rec in this.updates)
                {
                    formulas[i] = rec.Formula;
                    i++;
                }

                var flags = this.FormulaFlags;
                
                int c = surface.SetFormulas(stream, formulas, (short) flags);
            }
        }
        
        private short [] build_stream()
        {
            var st = this.first_update.Value.StreamType;

            if (st==StreamType.SRC)
            {
                var streamb = new List<SRC>(this.updates.Count);
                streamb.AddRange( this.updates.Select(i=>i.SIDSRC.SRC));
                return SRC.ToStream(streamb);
            }
            else
            {
                var streamb = new List<SIDSRC>(this.updates.Count);
                streamb.AddRange(this.updates.Select(i => i.SIDSRC));
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

        public void SetResult(SRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(StreamType.SRC, new SIDSRC(-1, streamitem), value, unitcode);
        }

        public void SetResult(SRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(StreamType.SRC, new SIDSRC(-1, streamitem), value, unitcode);
        }

        public struct UpdateRecord
        {
            public readonly SIDSRC SIDSRC;
            public readonly string Formula;
            public readonly double ResultNumeric;
            public readonly string ResultString;
            public readonly IVisio.VisUnitCodes UnitCode;
            public readonly UpdateType UpdateType;
            public readonly StreamType StreamType;

            internal UpdateRecord(StreamType st, SIDSRC sidsrc, string formula)
            {
                this.SIDSRC = sidsrc;
                this.Formula = formula;
                this.ResultNumeric = 0.0;
                this.ResultString = null;
                this.UnitCode = IVisio.VisUnitCodes.visNumber;
                this.UpdateType = UpdateType.Formula;
                this.StreamType = st;
            }

            internal UpdateRecord(StreamType st, SIDSRC sidsrc, double result, IVisio.VisUnitCodes unit_code)
            {
                this.SIDSRC = sidsrc;
                this.Formula = null;
                this.UnitCode = unit_code;
                this.ResultNumeric = result;
                this.ResultString = null;
                this.UpdateType = UpdateType.ResultNumeric;
                this.StreamType = st;
            }

            internal UpdateRecord(StreamType st, SIDSRC sidsrc, string result, IVisio.VisUnitCodes unit_code)
            {
                this.SIDSRC = sidsrc;
                this.Formula = null;
                this.UnitCode = unit_code;
                this.ResultNumeric = 0.0;
                this.ResultString = result;
                this.UpdateType = UpdateType.ResultString;
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
            ResultNumeric,
            ResultString
        }
    }
}