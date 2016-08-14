using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Update
{
    public class UpdateSRCBase : UpdateBase<SRC>
    {
        public UpdateSRCBase() : base()
        {
        }

        public UpdateSRCBase(int capacity) : base(capacity)
        {
        }

        protected override short[] build_stream()
        {
            var streamb = new List<SRC>(this._updates.Count);
            streamb.AddRange(this._updates.Select(i => i.StreamItem));
            return SRC.ToStream(streamb);
        }
    }

    public class UpdateSRCFormulas : UpdateSRCBase
    {
        public UpdateSRCFormulas() :base()
        {
        }

        public UpdateSRCFormulas(int capacity) : base( capacity )
        {
        }

        public void SetFormula(SRC streamitem, FormulaLiteral formula)
        {
            this._SetFormula(streamitem, formula);
        }

        public void SetFormulaIgnoreNull(SRC streamitem, FormulaLiteral formula)
        {
            this._SetFormulaIgnoreNull(streamitem, formula);
        }

        protected void _SetFormula(SRC streamitem, FormulaLiteral formula)
        {
            this.CheckFormulaIsNotNull(formula.Value);
            var rec = new UpdateRecord<SRC>(StreamType.SRC, streamitem, formula.Value);
            this._add_update(rec);
        }
        protected void _SetFormulaIgnoreNull(SRC streamitem, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this._SetFormula(streamitem, formula);
            }
        }
    }


    public class UpdateSRCResult : UpdateSRCBase
    {
        public UpdateSRCResult() : base()
        {
        }

        public UpdateSRCResult(int capacity) : base(capacity)
        {
        }


        public void SetResult(SRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(SRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(streamitem, value, unitcode);
        }

        protected void _SetResult(SRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            var rec = new UpdateRecord<SRC>(StreamType.SRC, streamitem, value, unitcode);
            this._add_update(rec);
        }

        protected void _SetResult(SRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            var rec = new UpdateRecord<SRC>(StreamType.SRC, streamitem, value, unitcode);
            this._add_update(rec);
        }

    }

}