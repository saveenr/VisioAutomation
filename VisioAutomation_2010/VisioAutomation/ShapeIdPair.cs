using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation
{
    public struct ShapeIDPair
    {
        public readonly IVisio.Shape Shape;
        public readonly int ShapeID;

        public ShapeIDPair(IVisio.Shape shape)
        {
            this.Shape = shape;
            this.ShapeID = shape.ID16;
        }

        public ShapeIDPair(IVisio.Shape shape, int id)
        {
            this.Shape = shape;
            this.ShapeID = id;
        }
    }

}

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
    public static class ExtensionHelpers
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
