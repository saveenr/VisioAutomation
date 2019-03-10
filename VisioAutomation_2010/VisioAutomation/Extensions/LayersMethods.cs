using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class LayersMethods
    {
        public static IEnumerable<IVisio.Layer> ToEnumerable(this IVisio.Layers layers)
        {
            return ExtensionHelpers.ToEnumerable(() => layers.Count, i => layers[i + 1]);
        }

        public static List<IVisio.Layer> ToList(this IVisio.Layers layers)
        {
            return ExtensionHelpers.ToList(() => layers.Count, i => layers[i + 1]);
        }
    }
}
