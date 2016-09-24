using System.Collections;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Set
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Nouns.VisioPageCell)]
    public class Set_VisioPageCell: VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true,Position=0)] 
        public Hashtable Hashtable  { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Pages { get; set; }

        protected override void ProcessRecord()
        {
            var target_pages = this.Pages ?? new[] { this.Client.Page.Get() };

            foreach (var page in target_pages)
            {
                var pagesheet = page.PageSheet;
                var t = new VisioAutomation.Scripting.TargetShapes(pagesheet);
                this.Client.ShapeSheet.SetPageCells( t , this.Hashtable, this.BlastGuards, this.TestCircular);
            }
        }
    }
}