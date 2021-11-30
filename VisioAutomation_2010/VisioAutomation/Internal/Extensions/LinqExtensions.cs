using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Internal.Extensions
{
    public static class LinqExtensions
    {
        public static IEnumerable<T> NotOfType<T>(this IEnumerable<T> enumerable, System.Type type)
        {
            return enumerable.Where(element => !type.IsAssignableFrom(element.GetType()));
        }
    }

}
