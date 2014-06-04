using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioRectangle")]
    public class New_VisioRectangle : VisioCmdlet
    {
        [System.Management.Automation.Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [System.Management.Automation.Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [System.Management.Automation.Parameter(Position = 2, Mandatory = true)]
        public double X1 { get; set; }

        [System.Management.Automation.Parameter(Position = 3, Mandatory = true)]
        public double Y1 { get; set; }

        protected override void ProcessRecord()
        {
            var rect = this.GetRectangle();
            var scriptingsession = this.ScriptingSession;
            var shape = scriptingsession.Draw.Rectangle(rect.Left,rect.Bottom,rect.Right,rect.Top);
            this.WriteObject(shape);
        }

        protected VisioAutomation.Drawing.Rectangle GetRectangle()
        {
            return new VisioAutomation.Drawing.Rectangle(X0, Y0, X1, Y1);
        }
    }
}