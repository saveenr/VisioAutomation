using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Selection
{
    public static class SelectionHelper
    {
        public static IList<IVisio.Shape> GetSelectedShapes(IVisio.Selection selection)
        {
            if (selection.Count < 1)
            {
                return new List<IVisio.Shape>(0);
            }
            
            var sel_shapes = selection.AsEnumerable();
            var shapes = sel_shapes.ToList();
            return shapes;
        }

        public static IList<IVisio.Shape> GetSelectedShapesRecursive(IVisio.Selection selection)
        {
            if (selection.Count < 1)
            {
                return new List<IVisio.Shape>(0);
            }

            var shapes = new List<IVisio.Shape>();
            var sel_shapes = selection.AsEnumerable();
            foreach (var shape in Shapes.ShapeHelper.GetNestedShapes(sel_shapes))
            {
                if (shape.Type != (short)IVisio.VisShapeTypes.visTypeGroup)
                {
                    shapes.Add(shape);
                }
            }
            return shapes;
        }

        public static void SendShapes(IVisio.Selection selection, ShapeSendDirection dir)
        {

            if (dir == ShapeSendDirection.ToBack)
            {
                selection.SendToBack();
            }
            else if (dir == ShapeSendDirection.Backward)
            {
                selection.SendBackward();
            }
            else if (dir == ShapeSendDirection.Forward)
            {
                selection.BringForward();
            }
            else if (dir == ShapeSendDirection.ToFront)
            {
                selection.BringToFront();
            }
        }
    }
}