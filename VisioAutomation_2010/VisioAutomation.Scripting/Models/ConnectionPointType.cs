using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    public enum ConnectionPointType
    {
        Inward = IVisio.VisCellVals.visCnnctTypeInward,
        Outward = IVisio.VisCellVals.visCnnctTypeOutward,
        InwardOutward = IVisio.VisCellVals.visCnnctTypeInwardOutward
    }
}