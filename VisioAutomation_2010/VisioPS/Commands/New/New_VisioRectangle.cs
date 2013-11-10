using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    public class RectangleCmdlet : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "NumberArray")]
        public double[] Doubles { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "SeparateNumbers")]
        public double X0 { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true, ParameterSetName = "SeparateNumbers")]
        public double Y0 { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = true, ParameterSetName = "SeparateNumbers")]
        public double X1 { get; set; }

        [SMA.Parameter(Position = 3, Mandatory = true, ParameterSetName = "SeparateNumbers")]
        public double Y1 { get; set; }

        protected VisioAutomation.Drawing.Rectangle GetRectangle()
        {
            if (this.Doubles != null)
            {
                if (Doubles.Length != 4)
                {
                    string msg = "Must have four vales";
                    throw new System.ArgumentOutOfRangeException(msg);
                }
                return new VisioAutomation.Drawing.Rectangle(Doubles[0], Doubles[1], Doubles[2], Doubles[3]);
            }
            else
            {
                return new VisioAutomation.Drawing.Rectangle(X0, Y0, X1, Y1);
            }
        }
    }

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