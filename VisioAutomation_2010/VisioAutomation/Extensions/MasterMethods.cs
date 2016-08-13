using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods
    {
        public static Drawing.Rectangle GetBoundingBox(this Microsoft.Office.Interop.Visio.Master master, Microsoft.Office.Interop.Visio.VisBoundingBoxArgs args)
        {
            return VisioAutomation.Masters.MasterHelper.GetBoundingBox(master, args);
        }

        public static IEnumerable<Master> ToEnumerable(this Microsoft.Office.Interop.Visio.Masters masters)
        {
            return VisioAutomation.Masters.MasterHelper.ToEnumerable(masters);
        }

        public static string[] GetNamesU(this Microsoft.Office.Interop.Visio.Masters masters)
        {
            return VisioAutomation.Masters.MasterHelper.GetNamesU(masters);
        }
    }
}