using SMA = System.Management.Automation;
using VisioAutomation.Shapes;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioHyperlink)]
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

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var hlink = new HyperlinkCells();

            hlink.Address = this.Address;
            hlink.Description= this.Description;
            hlink.ExtraInfo= this.ExtraInfo;
            hlink.Frame= this.Frame;
            hlink.SortKey= this.SortKey;

            hlink.SubAddress = this.SubAddress;

            hlink.Default = this.Default;
            hlink.NewWindow = this.NewWindow;
            hlink.Invisible = this.Invisible;

            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            this.Client.Hyperlink.Add(targets, hlink);
        }
    }
}