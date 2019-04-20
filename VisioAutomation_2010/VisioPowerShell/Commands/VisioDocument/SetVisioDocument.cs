using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioDocument
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioDocument)]
    public class SetVisioDocument : VisioCmdlet
    {

        // NONCONTEXT:DOCUMENT
        [SMA.Parameter(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNull]
        public IVisio.Document Document  { get; set; }
        
        protected override void ProcessRecord()
        {
            this.Client.Document.ActivateDocument(this.Document);
        }
    }
}