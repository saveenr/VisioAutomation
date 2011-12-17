using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("New", "ShapeSheetUpdate")]
    public class New_ShapeSheetUpdate : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var update = new VA.Scripting.ShapeSheetUpdate();
            this.WriteObject(update);
        }
    }
}