using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioShape")]
    public class New_VisioShape : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public IVisio.Master[] Masters { get; set; }

        [SMA.ParameterAttribute(Position = 1, Mandatory = true)]
        public double [] Points { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter NoSelect=false;

        protected override void ProcessRecord()
        {
            this.WriteVerbose("NoSelect: {0}", this.NoSelect);

            var points = VisioAutomation.Drawing.Point.FromDoubles(this.Points).ToList();
            var shape_ids = this.client.Master.Drop(this.Masters, points);

            var page = this.client.Page.Get();
            var shape_objects = VisioAutomation.Shapes.ShapeHelper.GetShapesFromIDs(page.Shapes, shape_ids);

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