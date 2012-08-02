using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Update
{
    public class SIDSRCUpdate : UpdateBase
    {
        public SIDSRCUpdate() :
            base()
        {
        }

        public SIDSRCUpdate(int capacity) :
            base(capacity)
        {
        }
        
        public void SetResult(short shapeid, SRC src, double value, IVisio.VisUnitCodes unitcode)
        {
            var streamitem = new SIDSRC(shapeid,src);
            this._SetResult(VA.ShapeSheet.Update.StreamType.SIDSRC, streamitem, value, unitcode);
        }

        public void SetResult(SIDSRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(VA.ShapeSheet.Update.StreamType.SIDSRC, streamitem, value, unitcode);
        }

        public void SetFormula(SIDSRC streamitem, FormulaLiteral formula)
        {
            this._SetFormula(VA.ShapeSheet.Update.StreamType.SIDSRC, streamitem, formula);
        }

        public void SetFormula(short shapeid, SRC src, FormulaLiteral formula)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this._SetFormula(VA.ShapeSheet.Update.StreamType.SIDSRC, streamitem, formula);
        }

        public void SetFormulaIgnoreNull(short id, ShapeSheet.SRC src, ShapeSheet.FormulaLiteral formula)
        {
            var sidsrc = new VA.ShapeSheet.SIDSRC(id, src);
            this._SetFormulaIgnoreNull(VA.ShapeSheet.Update.StreamType.SIDSRC, sidsrc, formula);
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

            var stream = GetResultStream();
            var unitcodes = this.GetUnitCodesArray();
            double[] results = this.GetResultsArray();
            var flags = this.ResultFlags;

            return VA.ShapeSheet.ShapeSheetHelper.SetResults(page, stream, results, unitcodes, flags);
        }

        private short [] GetResultStream()
        {
            var stream = new List<SIDSRC>(this.ResultCount);
            stream.AddRange(this.ResultRecords.Select(i => i.SIDSRC));
            return SIDSRC.ToStream(stream);
        }

        private short SetFormulas(IVisio.Page page)
        {
            if (this.FormulaCount == 0)
            {
                return 0;
            }

            var stream = GetFormulaStream();
            var formulas = this.GetFormulasArray();
            var flags = this.FormulaFlags;

            return VA.ShapeSheet.ShapeSheetHelper.SetFormulas(page, stream, formulas, (short)flags);
        }

        private short [] GetFormulaStream()
        {
            var stream = new List<SIDSRC>(this.FormulaCount);
            stream.AddRange(this.FormulaRecords.Select(i => i.SIDSRC));
            return SIDSRC.ToStream(stream);
        }
    }
}