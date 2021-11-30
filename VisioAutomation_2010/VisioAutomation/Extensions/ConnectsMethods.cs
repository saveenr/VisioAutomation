using System.Collections.Generic;
using VisioAutomation.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ConnectsMethods
    {
        public static IEnumerable<IVisio.Connect> ToEnumerable(this IVisio.Connects connects)
        {
            return CollectionHelpers.ToEnumerable(() => connects.Count, i => connects[i + 1]);
        }

        public static List<IVisio.Connect> ToList(this IVisio.Connects connects)
        {
            return CollectionHelpers.ToList(() => connects.Count, i => connects[i + 1]);
        }
    }
}
