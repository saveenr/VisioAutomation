using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IList<IVisio.Shape> GetShapesFromIDs(IList<short> shapeids)
        {
            var page = this.ActivePage;
            var page_shapes = page.Shapes;
            var shapes = VA.ShapeHelper.GetShapesFromIDs(page_shapes, shapeids);
            return shapes;
        }

        public void SelectAll()
        {
            this.ScriptingSession.Selection.SelectAll();
        }

        public void SelectNone()
        {
            this.ScriptingSession.Selection.SelectNone();
        }

        public void InvertSelection()
        {
            this.ScriptingSession.Selection.SelectInvert();
        }

        public void Select(IVisio.Shape shape)
        {
            this.ScriptingSession.Selection.SelectShape(shape);
        }

        public void Select(IList<IVisio.Shape> shapes)
        {
            this.ScriptingSession.Selection.SelectShapes(shapes);
        }

        public void SubSelect(IList<IVisio.Shape> shapes)
        {
            this.ScriptingSession.Selection.SubSelect(shapes);
        }

        public void SelectByLayer(string layername)
        {
            this.ScriptingSession.Selection.SelectShapesInLayer(layername);
        }

        public void SelectByMaster(IVisio.Master master)
        {
            this.ScriptingSession.Selection.SelectShapesByMaster(master);
        }
        
        public IList<IVisio.Shape> GetSelectedShapes()
        {
            var ss = this.ScriptingSession;
            return ss.Connection.GetSelectedShapes(VisioAutomation.ShapesEnumeration.Flat);
        }

    }
}