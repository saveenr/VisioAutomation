using System.Collections.Generic;
using VisioAutomation.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class SectionMethods
    {
        public static IEnumerable<IVisio.Row> ToEnumerable(this IVisio.Section section)
        {
            // Section object: http://msdn.microsoft.com/en-us/library/ms408988(v=office.12).aspx
            return CollectionHelpers.ToEnumerable(() => section.Count, i => section[(short)i]);
        }

        public static List<IVisio.Row> ToList(this IVisio.Section section)
        {
            return CollectionHelpers.ToList(() => section.Count, i => section[(short)i]);
        }
    }
}