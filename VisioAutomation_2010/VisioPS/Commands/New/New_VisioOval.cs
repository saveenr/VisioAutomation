using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioOval")]
    public class New_VisioOval : RectangleCmdlet
    {

        protected override void ProcessRecord()
        {
            var rect = this.GetRectangle();
            var scriptingsession = this.ScriptingSession;
            var shape = scriptingsession.Draw.Oval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            this.WriteObject(shape);
        }
    }
}