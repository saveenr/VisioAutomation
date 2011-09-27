using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Update
{
    public class SRCUpdate : UpdateBase<SRC>
    {
        public SRCUpdate() :
            base()
        {
        }

        public SRCUpdate(int fcapacity,int rcapacity)
            :base(fcapacity,rcapacity)
        {
        }
        
        public void Execute(IVisio.Shape shape)
        {
            this.SetResults(shape);
            this.SetFormulas(shape);
        }

        private short SetResults(
            IVisio.Shape shape)
        {
            if (this.ResultData.Count== 0)
            {
                return 0;
            }

            var stream = VA.ShapeSheet.Streams.SRCStream.FromItems(this.ResultData.Items, r => r.StreamItem);
            var unitcodes = this.ResultData.GetUnitCodesArray();
            var results = this.ResultData.GetResultsArray();
            var flags = this.ResultFlags;
            return VA.ShapeSheet.ShapeSheetHelper.SetResults(shape, stream, results, unitcodes, flags);
        }

        private short SetFormulas(IVisio.Shape shape)
        {
            if (this.FormulaData.Count == 0)
            {
                return 0;
            }

            var stream = VA.ShapeSheet.Streams.SRCStream.FromItems(this.FormulaData.Items, r => r.StreamItem);
            var formulas = this.FormulaData.GetFormulasArray();

            var flags = this.FormulaFlags;
            return VA.ShapeSheet.ShapeSheetHelper.SetFormulas(shape, stream, formulas, flags);
        }
    }
}