using System.Collections.Generic;

namespace VisioAutomation.Internal.Extensions
{
    internal static class ExtensionHelpers
    {

        public static IEnumerable<T> ToEnumerable<T>(System.Func<int> get_count, System.Func<int, T> get_item)
        {
            int count = get_count();
            for (int i = 0; i < count; i++)
            {
                var item = get_item(i);
                yield return item;
            }
        }

        public static List<T> ToList<T>(System.Func<int> get_count, System.Func<int, T> get_item)
        {
            int count = get_count();
            var list = new List<T>(count);
            for (int i = 0; i < count; i++)
            {
                var item = get_item(i);
                list.Add(item);
            }

            return list;
        }


    }
}
