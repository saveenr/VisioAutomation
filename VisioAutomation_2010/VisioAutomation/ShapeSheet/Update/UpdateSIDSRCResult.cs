using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Update
{
    public class UpdateSIDSRCBase : UpdateBase<SIDSRC>
    {
        public UpdateSIDSRCBase() : base()
        {
        }

        public UpdateSIDSRCBase(int capacity) : base(capacity)
        {
        }

        protected override short[] build_stream()
        {
            var streamb = new List<SIDSRC>(this._updates.Count);
            streamb.AddRange(this._updates.Select(i => i.StreamItem));
            return SIDSRC.ToStream(streamb);
        }
    }

    public class UpdateSIDSRCResult : UpdateSIDSRCBase
    {

        public UpdateSIDSRCResult() : base()
        {
        }

        public UpdateSIDSRCResult(int capacity) : base(capacity)
        {
        }


        public void SetResult(short shapeid, SRC src, double value, IVisio.VisUnitCodes unitcode)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(SIDSRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(short shapeid, SRC src, string value, IVisio.VisUnitCodes unitcode)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(SIDSRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(streamitem, value, unitcode);
        }


        protected void _SetResult(SIDSRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            var rec = new UpdateRecord<SIDSRC>(StreamType.SIDSRC, streamitem, value, unitcode);
            this._add_update(rec);
        }

        protected void _SetResult(SIDSRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            var rec = new UpdateRecord<SIDSRC>(StreamType.SIDSRC, streamitem, value, unitcode);
            this._add_update(rec);
        }
    }

    public class UpdateSIDSRCFormula : UpdateSIDSRCBase
    {
        public UpdateSIDSRCFormula() : base()
        {
        }

        public UpdateSIDSRCFormula(int capacity) : base(capacity)
        {
        }

        protected void _SetFormula(SIDSRC streamitem, FormulaLiteral formula)
        {
            this.CheckFormulaIsNotNull(formula.Value);
            var rec = new UpdateRecord<SIDSRC>(StreamType.SIDSRC, streamitem, formula.Value);
            this._add_update(rec);
        }

        public void SetFormula(SIDSRC streamitem, FormulaLiteral formula)
        {
            this._SetFormula(streamitem, formula);
        }

        public void SetFormula(short shapeid, SRC src, FormulaLiteral formula)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this._SetFormula(streamitem, formula);
        }

        public void SetFormulaIgnoreNull(SIDSRC streamitem, FormulaLiteral formula)
        {
            this._SetFormulaIgnoreNull(streamitem, formula);
        }

        public void SetFormulaIgnoreNull(short id, SRC src, FormulaLiteral formula)
        {
            var sidsrc = new SIDSRC(id, src);
            this._SetFormulaIgnoreNull(sidsrc, formula);
        }

        protected void _SetFormulaIgnoreNull(SIDSRC streamitem, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this._SetFormula(streamitem, formula);
            }
        }

    }

}