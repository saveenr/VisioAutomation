using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Selections
{
    public static class SelectionHelper
    {
        public static IEnumerable<IVisio.Shape> ToEnumerable(IVisio.Selection selection)
        {
            short count16 = selection.Count16;
            for (short i = 0; i < count16; i++)
            {
                yield return selection[i + 1];
            }
        }

        public static Drawing.Rectangle GetBoundingBox(IVisio.Selection selection, IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;
            selection.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new Drawing.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static int[] GetIDs(IVisio.Selection selection)
        {
            System.Array ids_sa;
            selection.GetIDs(out ids_sa);
            int[] ids = (int[])ids_sa;
            return ids;
        }
        public static IList<IVisio.Shape> GetSelectedShapes(IVisio.Selection selection)
        {
            if (selection.Count < 1)
            {
                return new List<IVisio.Shape>(0);
            }
            
            var sel_shapes = selection.ToEnumerable();
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
            var sel_shapes = selection.ToEnumerable();
            foreach (var shape in Shapes.ShapeHelper.GetNestedShapes(sel_shapes))
            {
                if (shape.Type != (short)IVisio.VisShapeTypes.visTypeGroup)
                {
                    shapes.Add(shape);
                }
            }
            return shapes;
        }

        public static void SendShapes(IVisio.Selection selection, VisioAutomation.Selections.ShapeSendDirection dir)
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