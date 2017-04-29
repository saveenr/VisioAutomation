using System.Management.Automation;
using VisioAutomation.Shapes;
using VisioAutomation.Utilities;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioHyperlink)]
    public class NewVisioHyperlink : VisioCmdlet
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
            var hlink = new HyperlinkCells();

            hlink.Address = this.Address;
            hlink.Description= this.Description;
            hlink.ExtraInfo= this.ExtraInfo;
            hlink.Frame= this.Frame;
            hlink.SortKey= this.SortKey;

            hlink.SubAddress = this.SubAddress;

            hlink.Default = Convert.BoolToFormula(this.Default);
            hlink.NewWindow = Convert.BoolToFormula(this.NewWindow);
            hlink.Invisible = Convert.BoolToFormula(this.Invisible);

            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            this.Client.Hyperlink.Add(targets, hlink);
        }
    }
}