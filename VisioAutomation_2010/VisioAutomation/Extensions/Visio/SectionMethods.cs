using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class SectionMethods
    {
        // Section object: http://msdn.microsoft.com/en-us/library/ms408988(v=office.12).aspx

        public static IEnumerable<IVisio.Row> AsEnumerable(this IVisio.Section section)
        {
            int row_count = section.Count;

            for (int i = 0; i < row_count; i++)
            {
                var row = section[(short)i];
                yield return row;
            }
        }
    }
}