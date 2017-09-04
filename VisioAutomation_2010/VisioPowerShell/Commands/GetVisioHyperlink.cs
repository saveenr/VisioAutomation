using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioHyperlink)]
    public class GetVisioHyperlink : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [Parameter(Mandatory = false)]
        public SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            var dic = this.Client.Hyperlink.Get(targets);

            if (this.GetCells)
            {
                this.WriteObject(dic);
                return;
            }

            foreach (var shape_points in dic)
            {
                var shape = shape_points.Key;
                var hyperlinks = shape_points.Value;
                int shapeid = shape.ID;

                foreach (var hyperlink in hyperlinks)
                {
                    var hl_formulas = new VisioPowerShell.Models.Hyperlink();

                    hl_formulas.ShapeID = shapeid;

                    hl_formulas.Address = hyperlink.Address.Formula;
                    hl_formulas.Default = hyperlink.Default.Formula;
                    hl_formulas.Description = hyperlink.Description.Formula;
                    hl_formulas.ExtraInfo = hyperlink.ExtraInfo.Formula;
                    hl_formulas.Frame = hyperlink.Frame.Formula;
                    hl_formulas.Invisible = hyperlink.Invisible.Formula;
                    hl_formulas.NewWindow = hyperlink.NewWindow.Formula;
                    hl_formulas.SortKey = hyperlink.SortKey.Formula;
                    hl_formulas.SubAddress = hyperlink.SubAddress.Formula;

                    this.WriteObject(hl_formulas);
                }
            }
        }
    }
}
 