using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Shapes.Connections
{
    public enum ConnectionPointType
    {
        Inward = IVisio.VisCellVals.visCnnctTypeInward,
        Outward = IVisio.VisCellVals.visCnnctTypeOutward,
        InwardOutward = IVisio.VisCellVals.visCnnctTypeInwardOutward
    }
}