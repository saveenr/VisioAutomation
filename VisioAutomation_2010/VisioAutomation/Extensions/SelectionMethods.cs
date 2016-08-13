using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class SelectionMethods
    {
        public static IEnumerable<Shape> ToEnumerable(this Microsoft.Office.Interop.Visio.Selection selection)
        {
            return VisioAutomation.Selections.SelectionHelper.ToEnumerable(selection);
        }
        
        public static Drawing.Rectangle GetBoundingBox(this Microsoft.Office.Interop.Visio.Selection selection, Microsoft.Office.Interop.Visio.VisBoundingBoxArgs args)
        {
            return VisioAutomation.Selections.SelectionHelper.GetBoundingBox(selection, args);
        }

        public static int[] GetIDs(this Microsoft.Office.Interop.Visio.Selection selection)
        {
            return VisioAutomation.Selections.SelectionHelper.GetIDs(selection);
        }
    }
}