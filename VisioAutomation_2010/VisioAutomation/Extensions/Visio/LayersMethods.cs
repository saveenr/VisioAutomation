using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class LayersMethods
    {
        public static IEnumerable<IVisio.Layer> AsEnumerable(this IVisio.Layers layers)
        {
            short count = layers.Count;
            for (int i = 0; i < count; i++)
            {
                yield return layers[i + 1];
            }
        }
    }
}