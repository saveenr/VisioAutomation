using VisioAutomation.ShapeSheet.Streams;
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
            return VA.ShapeSheet.ShapeSheetHelper.SetResults(shape, stream, results, unitcodes, flags);
        }

        private SRCStream GetResultStream()
        {
            var stream = new VA.ShapeSheet.Streams.SRCStream(this.ResultCount);
            stream.AddRange(this.ResultRecords.Select(i => i.StreamItem));
            return stream;
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

        private SRCStream GetFormulaStream()
        {
            var stream = new VA.ShapeSheet.Streams.SRCStream(this.FormulaCount);
            stream.AddRange(this.FormulaRecords.Select(i => i.StreamItem));
            return stream;
        }
    }
}