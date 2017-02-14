using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Scripting.ShapeSheet
{
    public class ShapeSheetReader
    {
        public Client Client;
        public VisioAutomation.ShapeSheet.ShapeSheetSurface Surface;
        public SidsrcShapeSheetStreamBuilder SidsrcShapeSheetStreamBuilder;
        
        public ShapeSheetReader(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.Surface = new ShapeSheetSurface(page);
            this.SidsrcShapeSheetStreamBuilder = new SidsrcShapeSheetStreamBuilder();
        }

        public void AddCell(short id, VisioAutomation.ShapeSheet.SRC src)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC(id, src);
            this.SidsrcShapeSheetStreamBuilder.Add(sidsrc);
        }

        public string[] GetFormulas()
        {
            var formulas = this.Surface.GetFormulasU(this.SidsrcShapeSheetStreamBuilder.ToStream());
            return formulas;
        }

        public string[] GetResults()
        {
            var unitcodes = new UnitCodesBuilder(1);
            unitcodes.Add(IVisio.VisUnitCodes.visNoCast);
            var formulas = this.Surface.GetResults<string>( this.SidsrcShapeSheetStreamBuilder.ToStream(), unitcodes);
            return formulas;
        }
    }
}