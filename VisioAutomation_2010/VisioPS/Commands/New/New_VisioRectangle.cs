using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioRectangle")]
    public class New_VisioRectangle : RectangleCmdlet
    {

        protected override void ProcessRecord()
        {
            var rect = this.GetRectangle();
            var scriptingsession = this.ScriptingSession;
            var shape = scriptingsession.Draw.Rectangle(rect.Left,rect.Bottom,rect.Right,rect.Top);
            this.WriteObject(shape);
        }
    }
}