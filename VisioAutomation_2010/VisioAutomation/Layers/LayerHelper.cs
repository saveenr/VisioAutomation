using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layers
{
    public static class LayerHelper
    {
        public static IEnumerable<IVisio.Layer> ToEnumerable(IVisio.Layers layers)
        {
            short count = layers.Count;
            for (int i = 0; i < count; i++)
            {
                yield return layers[i + 1];
            }
        }
    }
}

