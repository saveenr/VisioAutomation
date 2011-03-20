using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class LayersMethods
    {
        public static IEnumerable<IVisio.Layer> AsEnumerable(this IVisio.Layers layers)
        {
            return layers.Cast<IVisio.Layer>();
        }
    }
}