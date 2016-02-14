using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Nouns.VisioHyperlink)]
    public class Get_VisioHyperlink : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [Parameter(Mandatory = false)]
        public SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var dic = this.Client.Hyperlink.Get(this.Shapes);

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
                    var hl_formulas = new HyperlinkFormulas();

                    hl_formulas.ShapeID = shapeid;

                    hl_formulas.Address = hyperlink.Address.Formula.Value;
                    hl_formulas.Default = hyperlink.Default.Formula.Value;
                    hl_formulas.Description = hyperlink.Description.Formula.Value;
                    hl_formulas.ExtraInfo = hyperlink.ExtraInfo.Formula.Value;
                    hl_formulas.Frame = hyperlink.Frame.Formula.Value;
                    hl_formulas.Invisible = hyperlink.Invisible.Formula.Value;
                    hl_formulas.NewWindow = hyperlink.NewWindow.Formula.Value;
                    hl_formulas.SortKey = hyperlink.SortKey.Formula.Value;
                    hl_formulas.SubAddress = hyperlink.SubAddress.Formula.Value;

                    this.WriteObject(hl_formulas);
                }
            }
        }
    }

    public class HyperlinkFormulas
    {
        public int ShapeID;
        public string Address { get; set; }
        public string Default { get; set; }
        public string Description { get; set; }
        public string ExtraInfo { get; set; }
        public string Frame { get; set; }
        public string Invisible { get; set; }
        public string NewWindow { get; set; }
        public string SortKey { get; set; }
        public string SubAddress { get; set; }

    }
}
 