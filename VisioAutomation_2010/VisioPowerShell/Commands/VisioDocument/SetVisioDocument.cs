using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioDocument)]
    public class SetVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Name")]
        [SMA.ValidateNotNullOrEmpty]
        public string Name { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Doc")]
        [SMA.ValidateNotNull]
        public IVisio.Document Document  { get; set; }
        
        protected override void ProcessRecord()
        {
            if (this.Name != null)
            {
                this.Client.Document.ActivateDocumentWithName(this.Name);
            }
            else if (this.Document != null)
            {
                this.Client.Document.ActivateDocument(this.Document);
            }
        }
    }
}