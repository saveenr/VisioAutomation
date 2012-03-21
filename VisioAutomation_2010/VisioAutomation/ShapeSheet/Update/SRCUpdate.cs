using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Update
{
    public class SRCUpdate : UpdateBase
    {
        public SRCUpdate() :
            base()
        {
        }

        public SRCUpdate(int capacity) :
            base(capacity)
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
            if (this.ResultCount== 0)
            {
                return 0;
            }

            var stream = GetResultStream();
            var unitcodes = this.GetUnitCodesArray();
            var results = this.GetResultsArray();
            var flags = this.ResultFlags;
            return VA.ShapeSheet.ShapeSheetHelper.SetResults(shape, stream, results, unitcodes, flags);
        }
         
        public void SetFormula(SRC streamitem, FormulaLiteral formula)
        {
            this._SetFormula(new SIDSRC(-1,streamitem), formula);
        }

        public void SetFormulaIgnoreNull(SRC streamitem, ShapeSheet.FormulaLiteral formula)
        {
            this._SetFormulaIgnoreNull(new SIDSRC(-1,streamitem),formula );
        }

        public void SetResult(SRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(new SIDSRC(-1,streamitem), value, unitcode);
        }

        private short [] GetResultStream()
        {
            var stream = new List<SRC>(this.ResultCount);
            stream.AddRange(this.ResultRecords.Select(i => i.SIDSRC.SRC));
            return SRC.ToStream(stream);
        }

        private short SetFormulas(IVisio.Shape shape)
        {
            if (this.FormulaCount == 0)
            {
                return 0;
            }

            var stream = GetFormulaStream();
            var formulas = this.GetFormulasArray();
            var flags = this.FormulaFlags;
            return VA.ShapeSheet.ShapeSheetHelper.SetFormulas(shape, stream, formulas, flags);
        }

        private short [] GetFormulaStream()
        {
            var stream = new List<SRC>(this.FormulaCount);
            stream.AddRange(this.FormulaRecords.Select(i => i.SIDSRC.SRC));
            return SRC.ToStream(stream);
        }
    }
}