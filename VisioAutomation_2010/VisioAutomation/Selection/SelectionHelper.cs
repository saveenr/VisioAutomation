using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Selection
{
    public static class SelectionHelper
    {
        public static IList<IVisio.Shape> GetSelectedShapes(IVisio.Selection selection, ShapesEnumeration enumerationtype)
        {
            if (selection.Count < 1)
            {
                return new List<IVisio.Shape>(0);
            }

            var shapes = selection.AsEnumerable();

            if (enumerationtype == ShapesEnumeration.Flat)
            {
                return shapes.ToList();
            }
            
            if (enumerationtype == ShapesEnumeration.ExpandGroups)
            {
                var shapes_in_groups = VA.ShapeHelper.GetNestedShapes(shapes)
                    .Where(s => s.Type != (short) IVisio.VisShapeTypes.visTypeGroup)
                    .ToList();
                return shapes_in_groups;
            }

            throw new System.ArgumentOutOfRangeException("enumerationtype");
        }

        public static void SendShapes(IVisio.Selection selection, VA.Selection.ShapeSendDirection dir)
        {

            if (dir == VA.Selection.ShapeSendDirection.ToBack)
            {
                selection.SendToBack();
            }
            else if (dir == VA.Selection.ShapeSendDirection.Backward)
            {
                selection.SendBackward();
            }
            else if (dir == VA.Selection.ShapeSendDirection.Forward)
            {
                selection.BringForward();
            }
            else if (dir == VA.Selection.ShapeSendDirection.ToFront)
            {
                selection.BringToFront();
            }
        }
    }
}