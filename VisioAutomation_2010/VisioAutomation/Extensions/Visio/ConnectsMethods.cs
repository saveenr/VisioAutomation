using System.Collections.Generic;
using VisioAutomation.Shapes.Connectors;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ConnectsMethods
    {
        public static IEnumerable<IVisio.Connect> ToEnumerable(this IVisio.Connects connects)
        {
            return ConnectorHelper.ToEnumerable(connects);
        }
    }
}
