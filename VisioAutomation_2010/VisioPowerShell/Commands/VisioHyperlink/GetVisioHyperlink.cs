using VisioAutomation.ShapeSheet;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioHyperlink
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
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);
            var dicof_shape_to_hyperlinks = this.Client.Hyperlink.GetHyperlinks(targetshapes, CellValueType.Formula);

            if (this.GetCells)
            {
                this.WriteObject(dicof_shape_to_hyperlinks);
                return;
            }

            foreach (var shape_hyperlinks_pair in dicof_shape_to_hyperlinks)
            {
                var shape = shape_hyperlinks_pair.Key;
                var hyperlinks = shape_hyperlinks_pair.Value;
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
 