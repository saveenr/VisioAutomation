using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Internal.Extensions
{
    public static class LinqExtensions
    {
        public static IEnumerable<T> WhereOfType<T>(this IEnumerable<T> enumerable, System.Type type)
        {
            return enumerable.Where(element => type.IsAssignableFrom(element.GetType()));
        }

        public static IEnumerable<T> WhereNotOfType<T>(this IEnumerable<T> enumerable, System.Type type)
        {
            return enumerable.Where(element => !type.IsAssignableFrom(element.GetType()));
        }
    }

}
