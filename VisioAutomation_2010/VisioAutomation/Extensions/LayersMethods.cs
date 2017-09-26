using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class LayersMethods
    {
        public static IEnumerable<IVisio.Layer> ToEnumerable(this IVisio.Layers layers)
        {
            short count = layers.Count;
            for (int i = 0; i < count; i++)
            {
                yield return layers[i + 1];
            }
        }

        public static List<IVisio.Layer> ToList(this IVisio.Layers layers)
        {
            int count = layers.Count;
            var list = new List<IVisio.Layer>(count);
            for (int i = 0; i < count; i++)
            {
                list.Add(layers[i + 1]);
            }
            return list;
        }
    }
}
