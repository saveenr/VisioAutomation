using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Draw", "Grid")]
    public class Draw_Grid : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VA.Layout.Grid.GridLayout GridLayout{ get; set; }

        [SMA.Parameter(Position = 5, Mandatory = true)]
        public double X = 0.0;

        [SMA.Parameter(Position = 6, Mandatory = true)]
        public double Y = 0.0;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var grid_origin = new VA.Drawing.Point(this.X, this.Y);
            var cellspacing = new VA.Drawing.Size(0, 0);
            scriptingsession.Draw.DrawGrid(this.GridLayout, grid_origin, cellspacing);
        }
    }
}