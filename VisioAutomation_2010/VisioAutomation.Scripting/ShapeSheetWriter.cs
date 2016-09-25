using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    public class ShapeSheetWriter
    {
        internal readonly FormulaWriterSIDSRC formula_writer;
        public Client Client;
        public IVisio.Page TargetPage;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetWriter(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.TargetPage = page;
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
                this.formula_writer.Commit(this.TargetPage);
            }
        }
    }

    public class ShapeSheetReader
    {
        public Client Client;
        public IVisio.Page TargetPage;
        public List<VisioAutomation.ShapeSheet.SIDSRC> SIDSRCs;
        
        public ShapeSheetReader(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.TargetPage = page;
            this.SIDSRCs = new List<VisioAutomation.ShapeSheet.SIDSRC>();
        }

        public void AddCell(short id, VisioAutomation.ShapeSheet.SRC src)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC(id, src);
            this.SIDSRCs.Add(sidsrc);
        }

        public string[] GetFormulas()
        {
            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(this.TargetPage);
            var streambuilder = new VisioAutomation.ShapeSheet.Queries.Utilities.StreamBuilderSIDSRC(this.SIDSRCs.Count);
            foreach (var sidsrc in this.SIDSRCs)
            {
                streambuilder.Add(sidsrc.ShapeID,sidsrc.SRC);
            }
            if (!streambuilder.IsFull)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }
            var formulas = VisioAutomation.ShapeSheet.Queries.Utilities.QueryHelpers.GetFormulasU_SIDSRC(surface,
                streambuilder.Stream);

            return formulas;
        }

        public string[] GetResults()
        {
            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(this.TargetPage);
            var streambuilder = new VisioAutomation.ShapeSheet.Queries.Utilities.StreamBuilderSIDSRC(this.SIDSRCs.Count);
            foreach (var sidsrc in this.SIDSRCs)
            {
                streambuilder.Add(sidsrc.ShapeID, sidsrc.SRC);
            }
            if (!streambuilder.IsFull)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var unitcodes = new List<IVisio.VisUnitCodes> {IVisio.VisUnitCodes.visNoCast};
            var formulas = VisioAutomation.ShapeSheet.Queries.Utilities.QueryHelpers.GetResults_SIDSRC<string>(surface,
                streambuilder.Stream, unitcodes);

            return formulas;
        }

    }
}