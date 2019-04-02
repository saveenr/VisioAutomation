using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioMaster
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioMaster)]
    public class GetVisioMaster : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string[] Name;

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public IVisio.Document Document;

        protected override void ProcessRecord()
        {
            var target_doc = new VisioScripting.TargetDocument(this.Document);
            target_doc.Resolve(this.Client);
            
            bool master_specified = this.Name !=null;
            bool doc_specified = this.Document !=null;

            if (master_specified)
            {
                foreach (var name in this.Name)
                {
                    var masters = this.Client.Master.FindMastersInDocumentByName(target_doc, name);
                    this.WriteObject(masters, true);
                }
            }
            else
            {
                // master name is not specified
                if (doc_specified)
                {
                    this.WriteVerbose("Get all masters from specified document");
                    var masters = this.Client.Master.GetAllMastersInDocument(target_doc);
                    this.WriteObject(masters, true);                    
                }
                else
                {
                    this.WriteVerbose("Get all masters from active document");
                    var masters = this.Client.Master.GetAllMastersInDocument(target_doc);
                    this.WriteObject(masters, true);                   
                }
            }
        }
    }
}