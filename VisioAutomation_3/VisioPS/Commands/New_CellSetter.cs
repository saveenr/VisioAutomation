using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("New", "New_CellSetter")]
    public class New_CellSetter : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var setter = new VA.Scripting.CellSetter();
            this.WriteObject(setter);
        }
    }
}