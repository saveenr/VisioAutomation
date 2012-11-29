using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VA=VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("New", "VisioShape")]
    public class New_VisioShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Master[] Masters { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double [] Points { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var points = VA.Drawing.Point.FromDoubles(Points).ToList();
            var shape_ids = scriptingsession.Master.Drop(Masters, points);
            var page = scriptingsession.Page.Get();
            var shape_objects = VA.ShapeHelper.GetShapesFromIDs(page.Shapes, shape_ids);
            this.WriteObject(shape_objects);
        }
    }
}