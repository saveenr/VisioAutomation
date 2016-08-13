using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class LayersMethods
    {
        public static IEnumerable<Layer> ToEnumerable(this Microsoft.Office.Interop.Visio.Layers layers)
        {
            return VisioAutomation.Layers.LayerHelper.ToEnumerable(layers);
        }
    }
}
