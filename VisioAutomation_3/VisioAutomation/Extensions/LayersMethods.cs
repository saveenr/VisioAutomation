using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class LayersMethods
    {
        public static IEnumerable<IVisio.Layer> AsEnumerable(this IVisio.Layers layers)
        {
            for (int i = 0; i < layers.Count; i++)
            {
                yield return layers[i + 1];
            }
        }
    }
}