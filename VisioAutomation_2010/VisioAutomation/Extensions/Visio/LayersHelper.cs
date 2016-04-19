using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layers
{
    public static class LayersHelper
    {
        public static IEnumerable<Layer> ToEnumerable(Microsoft.Office.Interop.Visio.Layers layers)
        {
            short count = layers.Count;
            for (int i = 0; i < count; i++)
            {
                yield return layers[i + 1];
            }
        }
    }
}