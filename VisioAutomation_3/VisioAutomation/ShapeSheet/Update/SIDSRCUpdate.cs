using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Streams;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Update
{
    public class SIDSRCUpdate : UpdateBase<SIDSRC>
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
            this.SetResult(streamitem, value, unitcode);
        }

        public void SetFormula(short shapeid, SRC src, FormulaLiteral formula)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this.SetFormula(streamitem, formula);
        }

        public void SetFormulaIgnoreNull(short id, ShapeSheet.SRC src, ShapeSheet.FormulaLiteral f)
        {
            var sidsrc = new VA.ShapeSheet.SIDSRC(id, src);
            this.SetFormulaIgnoreNull(sidsrc,f);
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

            return VA.ShapeSheet.ShapeSheetHelper.SetResults(page, stream, results, unitcodes, flags, this.ResultCount);
        }

        private short [] GetResultStream()
        {
            var stream = new List<SIDSRC>(this.ResultCount);
            stream.AddRange(this.ResultRecords.Select(i => i.StreamItem));
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

            return VA.ShapeSheet.ShapeSheetHelper.SetFormulas(page, stream, formulas, (short)flags, this.FormulaCount);
        }

        private short [] GetFormulaStream()
        {
            var stream = new List<SIDSRC>(this.FormulaCount);
            stream.AddRange(this.FormulaRecords.Select(i => i.StreamItem));
            return SIDSRC.ToStream(stream);
        }
    }
}