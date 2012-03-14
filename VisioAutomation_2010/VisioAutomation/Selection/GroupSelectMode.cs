using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Selection
{
    public enum GroupSelectMode
    {
        GroupFirst = IVisio.VisCellVals.visGrpSelModeGroup1st,
        GroupOnly = IVisio.VisCellVals.visGrpSelModeGroupOnly,
        MembersFirst = IVisio.VisCellVals.visGrpSelModeMembers1st
    }
}