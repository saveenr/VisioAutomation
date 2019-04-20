using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioHyperlink
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioHyperlink)]
    public class NewVisioHyperlink : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true)]
        public string Address { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string Description { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string SubAddress { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ExtraInfo { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string Frame { get; set; }
 
        [SMA.Parameter(Mandatory = false)]
        public string SortKey { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public bool NewWindow { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public bool Default { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public bool Invisible { get; set; }

        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            var hlink = new VisioAutomation.Shapes.HyperlinkCells();

            hlink.Address = this.Address;
            hlink.Description= this.Description;
            hlink.ExtraInfo= this.ExtraInfo;
            hlink.Frame= this.Frame;
            hlink.SortKey= this.SortKey;

            hlink.SubAddress = this.SubAddress;

            hlink.Default = this.Default;
            hlink.NewWindow = this.NewWindow;
            hlink.Invisible = this.Invisible;

            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            this.Client.Hyperlink.AddHyperlink(targetshapes, hlink);
        }
    }
}