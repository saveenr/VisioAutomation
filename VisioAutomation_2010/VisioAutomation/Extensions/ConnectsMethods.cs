using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;
using VisioAutomation.Shapes.Connectors;

namespace VisioAutomation.Extensions
{
    public static class ConnectsMethods
    {
        public static IEnumerable<Connect> ToEnumerable(this Microsoft.Office.Interop.Visio.Connects connects)
        {
            return ConnectorHelper.ToEnumerable(connects);
        }
    }
}
