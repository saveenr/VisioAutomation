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
            return VA.ShapeSheet.Update.UpdateUtil.SetResults(shape, stream, results, unitcodes, flags, this.ResultCount);
        }

        private short [] GetResultStream()
        {
            var stream = new List<SRC>(this.ResultCount);
            stream.AddRange(this.ResultRecords.Select(i => i.StreamItem));
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
            return VA.ShapeSheet.Update.UpdateUtil.SetFormulas(shape, stream, formulas, flags, this.FormulaCount);
        }

        private short [] GetFormulaStream()
        {
            var stream = new List<SRC>(this.FormulaCount);
            stream.AddRange(this.FormulaRecords.Select(i => i.StreamItem));
            return SRC.ToStream(stream);
        }
    }
}