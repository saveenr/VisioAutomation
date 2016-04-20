using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Section
{
    public static class SectionHelper
    {
        public static IEnumerable<Row> ToEnumerable(Microsoft.Office.Interop.Visio.Section section)
        {
            // Section object: http://msdn.microsoft.com/en-us/library/ms408988(v=office.12).aspx

            int row_count = section.Count;

            for (int i = 0; i < row_count; i++)
            {
                var row = section[(short)i];
                yield return row;
            }
        }
    }
}