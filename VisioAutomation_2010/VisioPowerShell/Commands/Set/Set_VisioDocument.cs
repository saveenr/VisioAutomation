using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Set
{
    [Cmdlet(VerbsCommon.Set, "VisioDocument")]
    public class Set_VisioDocument : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true, ParameterSetName = "Name")]
        [ValidateNotNullOrEmpty]
        public string Name { get; set; }

        [Parameter(Position = 0, Mandatory = true, ParameterSetName = "Doc")]
        [ValidateNotNull]
        public IVisio.Document Document  { get; set; }
        
        protected override void ProcessRecord()
        {
            if (this.Name != null)
            {
                this.client.Document.Activate(this.Name);
            }
            else if (this.Document != null)
            {
                this.client.Document.Activate(this.Document);
            }
        }
    }
}