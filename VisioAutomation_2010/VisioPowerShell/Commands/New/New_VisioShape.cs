using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VA=VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioShape")]
    public class New_VisioShape : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Master[] Masters { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double [] Points { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter NoSelect=false;

        protected override void ProcessRecord()
        {
            this.WriteVerbose("NoSelect: {0}", this.NoSelect);

            var points = VA.Drawing.Point.FromDoubles(Points).ToList();
            var shape_ids = this.client.Master.Drop(Masters, points);

            var page = this.client.Page.Get();
            var shape_objects = VA.Shapes.ShapeHelper.GetShapesFromIDs(page.Shapes, shape_ids);

            this.client.Selection.None();

            if (this.NoSelect)
            {
            }
            else
            {
                ((SMA.Cmdlet) this).WriteVerbose("Selecting");
                this.client.Selection.Select(shape_objects);
            }

            this.WriteObject(shape_objects, false);
        }
    }
}