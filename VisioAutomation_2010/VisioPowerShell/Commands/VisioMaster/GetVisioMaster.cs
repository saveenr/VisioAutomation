using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioMaster
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioMaster)]
    public class GetVisioMaster : VisioCmdlet
    {
        //TODO: SHould this be a find cmdlet instead of a get cmdlet?

        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string[] Name;

        // NONCONTEXT:DOCUMENT
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document Document;

        protected override void ProcessRecord()
        {
            var targetdoc = new VisioScripting.TargetDocument(this.Document);
            targetdoc.Resolve(this.Client);
            
            bool master_specified = this.Name !=null;
            bool doc_specified = this.Document !=null;

            if (master_specified)
            {
                foreach (var name in this.Name)
                {
                    var masters = this.Client.Master.FindMasters(targetdoc, name);
                    this.WriteObject(masters, true);
                }
            }
            else
            {
                // master name is not specified
                if (doc_specified)
                {
                    this.WriteVerbose("Get all masters from specified document");
                    var masters = this.Client.Master.GetMasters(targetdoc);
                    this.WriteObject(masters, true);                    
                }
                else
                {
                    this.WriteVerbose("Get all masters from active document");
                    var masters = this.Client.Master.GetMasters(targetdoc);
                    this.WriteObject(masters, true);                   
                }
            }
        }
    }
}