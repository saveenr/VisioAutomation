using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Helpers
{
    public static class SelectionHelper
    {
        public static List<IVisio.Shape> GetSelectedShapes(IVisio.Selection selection)
        {
            if (selection.Count < 1)
            {
                return new List<IVisio.Shape>(0);
            }
            
            var sel_shapes = selection.ToEnumerable();
            var shapes = sel_shapes.ToList();
            return shapes;
        }

        public static List<IVisio.Shape> GetSelectedShapesRecursive(IVisio.Selection selection)
        {
            if (selection.Count < 1)
            {
                return new List<IVisio.Shape>(0);
            }

            var shapes = new List<IVisio.Shape>();
            var sel_shapes = selection.ToEnumerable();
            foreach (var shape in VisioAutomation.Shapes.ShapeHelper.GetNestedShapes(sel_shapes))
            {
                if (shape.Type != (short)IVisio.VisShapeTypes.visTypeGroup)
                {
                    shapes.Add(shape);
                }
            }
            return shapes;
        }
    }
}