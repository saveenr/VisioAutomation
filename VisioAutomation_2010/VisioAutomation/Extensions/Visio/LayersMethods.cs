using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class LayersMethods
    {
        public static IEnumerable<IVisio.Layer> ToEnumerable(this IVisio.Layers layers)
        {
            return VisioAutomation.Layers.LayersHelper.ToEnumerable(layers);
        }
    }
}
