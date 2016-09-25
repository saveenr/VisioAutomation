using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.Scripting.ShapeSheet
{
    public class ShapeSheetWriter
    {
        internal readonly FormulaWriterSIDSRC formula_writer;
        public Client Client;
        public VisioAutomation.ShapeSheet.ShapeSheetSurface Surface;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetWriter(Client client, Microsoft.Office.Interop.Visio.Page page)
        {
            this.Client = client;
            this.Surface= new ShapeSheetSurface(page);
            this.formula_writer = new FormulaWriterSIDSRC();
        }

        public void SetFormula(short id, VisioAutomation.ShapeSheet.SRC src, string formula)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC(id, src);
            this.formula_writer.SetFormula(sidsrc, formula);
        }

        public void Commit()
        {
            using (var undoscope = this.Client.Application.NewUndoScope("Modify ShapeSheet"))
            {
                this.formula_writer.BlastGuards = this.BlastGuards;
                this.formula_writer.TestCircular = this.TestCircular;
                this.formula_writer.Commit(this.Surface);
            }
        }
    }

    public class ShapeSheetReader
    {
        public Client Client;
        public VisioAutomation.ShapeSheet.ShapeSheetSurface Surface;
        public List<VisioAutomation.ShapeSheet.SIDSRC> SIDSRCs;
        
        public ShapeSheetReader(Client client, Microsoft.Office.Interop.Visio.Page page)
        {
            this.Client = client;
            this.Surface = new ShapeSheetSurface(page);
            this.SIDSRCs = new List<VisioAutomation.ShapeSheet.SIDSRC>();
        }

        public void AddCell(short id, VisioAutomation.ShapeSheet.SRC src)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC(id, src);
            this.SIDSRCs.Add(sidsrc);
        }

        public string[] GetFormulas()
        {
            var stream = get_Stream();
            var formulas = VisioAutomation.ShapeSheet.Queries.Utilities.QueryHelpers.GetFormulasU_SIDSRC(this.Surface, stream);

            return formulas;
        }

        private short[] get_Stream()
        {
            var streambuilder = new VisioAutomation.ShapeSheet.Queries.Utilities.StreamBuilderSIDSRC(this.SIDSRCs.Count);
            foreach (var sidsrc in this.SIDSRCs)
            {
                streambuilder.Add(sidsrc.ShapeID, sidsrc.SRC);
            }
            if (!streambuilder.IsFull)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var stream = streambuilder.Stream;
            return stream;
        }

        public string[] GetResults()
        {
            var stream = get_Stream();
            var unitcodes = new List<VisUnitCodes> {Microsoft.Office.Interop.Visio.VisUnitCodes.visNoCast};
            var formulas = VisioAutomation.ShapeSheet.Queries.Utilities.QueryHelpers.GetResults_SIDSRC<string>(this.Surface, stream, unitcodes);
            return formulas;
        }

    }
}