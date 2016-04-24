using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods
    {
        public static Drawing.Rectangle GetBoundingBox(this IVisio.Master master, IVisio.VisBoundingBoxArgs args)
        {
            return VisioAutomation.Masters.MasterHelper.GetBoundingBox(master, args);
        }

        public static IEnumerable<IVisio.Master> ToEnumerable(this IVisio.Masters masters)
        {
            return VisioAutomation.Masters.MasterHelper.ToEnumerable(masters);
        }

        public static string[] GetNamesU(this IVisio.Masters masters)
        {
            return VisioAutomation.Masters.MasterHelper.GetNamesU(masters);
        }
    }
}