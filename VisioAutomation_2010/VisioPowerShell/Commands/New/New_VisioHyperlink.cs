using System.Management.Automation;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Nouns.VisioHyperlink)]
    public class New_VisioHyperlink : VisioCmdlet
    {
        [Parameter(Mandatory = true)]
        public string Address { get; set; }

        [Parameter(Mandatory = false)]
        public string Description { get; set; }

        [Parameter(Mandatory = false)]
        public string SubAddress { get; set; }

        [Parameter(Mandatory = false)]
        public string ExtraInfo { get; set; }

        [Parameter(Mandatory = false)]
        public string Frame { get; set; }
 
        [Parameter(Mandatory = false)]
        public string SortKey { get; set; }

        [Parameter(Mandatory = false)]
        public bool NewWindow { get; set; }

        [Parameter(Mandatory = false)]
        public bool Default { get; set; }

        [Parameter(Mandatory = false)]
        public bool Invisible { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var hlink = new VA.Shapes.Hyperlinks.HyperlinkCells();

            hlink.Address = this.Address;
            hlink.Description= this.Description;
            hlink.ExtraInfo= this.ExtraInfo;
            hlink.Frame= this.Frame;
            hlink.SortKey= this.SortKey;

            hlink.SubAddress = this.SubAddress;

            hlink.Default = VA.Convert.BoolToFormula(this.Default);
            hlink.NewWindow = VA.Convert.BoolToFormula(this.NewWindow);
            hlink.Invisible = VA.Convert.BoolToFormula(this.Invisible);

            this.Client.Hyperlink.Add(this.Shapes, hlink);
        }
    }
}