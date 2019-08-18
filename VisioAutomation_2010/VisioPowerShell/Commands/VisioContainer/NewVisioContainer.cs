using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioContainer
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioContainer)]
    public class NewVisioContainer : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true)]
        [SMA.ValidateNotNull]
        public IVisio.Master Master { get; set; }

        protected override void ProcessRecord()
        {
            var shape = this.Client.Container.DropContainerMaster(VisioScripting.TargetPage.Auto, this.Master);
            this.WriteObject(shape);
        }
    }
}
