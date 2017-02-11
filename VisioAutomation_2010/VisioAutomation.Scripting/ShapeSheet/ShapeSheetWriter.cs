﻿namespace VisioAutomation.Scripting.ShapeSheet
{
    public class ShapeSheetWriter
    {
        internal readonly VisioAutomation.ShapeSheet.Writers.ShapeSheetWriter writer;
        public Client Client;
        public VisioAutomation.ShapeSheet.ShapeSheetSurface Surface;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetWriter(Client client, Microsoft.Office.Interop.Visio.Page page)
        {
            this.Client = client;
            this.Surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page);
            this.writer = new VisioAutomation.ShapeSheet.Writers.ShapeSheetWriter();
        }

        public void SetFormula(short id, VisioAutomation.ShapeSheet.SRC src, string formula)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC(id, src);
            this.writer.SetFormula(sidsrc, formula);
        }

        public void Commit()
        {
            using (var undoscope = this.Client.Application.NewUndoScope("Modify ShapeSheet"))
            {
                this.writer.BlastGuards = this.BlastGuards;
                this.writer.TestCircular = this.TestCircular;
                this.writer.Commit(this.Surface);
            }
        }
    }
}