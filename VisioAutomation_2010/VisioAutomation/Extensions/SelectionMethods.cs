using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class SelectionMethods
    {
        public static IEnumerable<IVisio.Shape> ToEnumerable(this IVisio.Selection selection)
        {
            return VisioAutomation.Selections.SelectionHelper.ToEnumerable(selection);
        }
        
        public static Drawing.Rectangle GetBoundingBox(this IVisio.Selection selection, IVisio.VisBoundingBoxArgs args)
        {
            return VisioAutomation.Selections.SelectionHelper.GetBoundingBox(selection, args);
        }

        public static int[] GetIDs(this IVisio.Selection selection)
        {
            return VisioAutomation.Selections.SelectionHelper.GetIDs(selection);
        }
    }
}