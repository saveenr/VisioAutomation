using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class SectionMethods
    {
        public static IEnumerable<Row> ToEnumerable(this Microsoft.Office.Interop.Visio.Section section)
        {
            return VisioAutomation.ShapeSheet.ShapeSheetHelper.ToEnumerable(section);
        }
    }
}