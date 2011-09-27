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

        public SIDSRCUpdate(int fcapacity,int rcapacity) :
            base(fcapacity,rcapacity)
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
            if (this.ResultData.Count== 0)
            {
                return 0;
            }

            var stream = VA.ShapeSheet.Streams.SIDSRCStream.FromItems(this.ResultData.Items, r => r.StreamItem);
            var unitcodes = this.ResultData.GetUnitCodesArray();
            double[] results = this.ResultData.GetResultsArray();
            var flags = this.ResultFlags;

            return VA.ShapeSheet.ShapeSheetHelper.SetResults(page, stream, results, unitcodes, flags);
        }

        private short SetFormulas(IVisio.Page page)
        {
            if (this.FormulaData.Count == 0)
            {
                return 0;
            }

            var stream = VA.ShapeSheet.Streams.SIDSRCStream.FromItems(this.FormulaData.Items, r => r.StreamItem);
            var formulas = this.FormulaData.GetFormulasArray();
            var flags = this.FormulaFlags;

            return VA.ShapeSheet.ShapeSheetHelper.SetFormulas(page, stream, formulas, (short)flags);
        }
    }
}