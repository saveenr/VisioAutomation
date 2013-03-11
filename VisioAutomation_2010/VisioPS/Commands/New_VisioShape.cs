using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VA=VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioShape")]
    public class New_VisioShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Master[] Masters { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double [] Points { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter NoSelect=false;

        protected override void ProcessRecord()
        {
            this.WriteVerboseEx("NoSelect: {0}", this.NoSelect);

            var scriptingsession = this.ScriptingSession;
            var points = VA.Drawing.Point.FromDoubles(Points).ToList();
            var shape_ids = scriptingsession.Master.Drop(Masters, points);
            
            var page = scriptingsession.Page.Get();
            var shape_objects = VA.ShapeHelper.GetShapesFromIDs(page.Shapes, shape_ids);

            scriptingsession.Selection.SelectNone();

            if (this.NoSelect)
            {
            }
            else
            {
                this.WriteVerbose("Selecting");
                scriptingsession.Selection.Select(shape_objects);
            }

            this.WriteObject(shape_objects);
        }
    }
}