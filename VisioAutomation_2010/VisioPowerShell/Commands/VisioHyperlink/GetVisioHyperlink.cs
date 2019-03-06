using VisioAutomation.ShapeSheet;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioHyperlink)]
    public class GetVisioHyperlink : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            var dic = this.Client.Hyperlink.GetHyperlinks(targets, CellValueType.Formula);

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

                    hl_formulas.Address = hyperlink.Address.Value;
                    hl_formulas.Default = hyperlink.Default.Value;
                    hl_formulas.Description = hyperlink.Description.Value;
                    hl_formulas.ExtraInfo = hyperlink.ExtraInfo.Value;
                    hl_formulas.Frame = hyperlink.Frame.Value;
                    hl_formulas.Invisible = hyperlink.Invisible.Value;
                    hl_formulas.NewWindow = hyperlink.NewWindow.Value;
                    hl_formulas.SortKey = hyperlink.SortKey.Value;
                    hl_formulas.SubAddress = hyperlink.SubAddress.Value;

                    this.WriteObject(hl_formulas);
                }
            }
        }
    }
}
 