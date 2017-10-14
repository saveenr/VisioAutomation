using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioMaster)]
    public class GetVisioMaster : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string[] Name;

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public IVisio.Document Document;

        protected override void ProcessRecord()
        {
            bool master_specified = this.Name !=null;
            bool doc_specified = this.Document !=null;

            if (master_specified)
            {
                // master name is specified
                if (doc_specified)
                {
                    ((SMA.Cmdlet) this).WriteVerbose("Get master from specified document");
                    foreach (var name in this.Name)
                    {
                        var masters = this.Client.Master.GetMastersByNameFromDocument(name, this.Document);
                        this.WriteObject(masters, true);
                    }
                }
                else
                {
                    ((SMA.Cmdlet) this).WriteVerbose("Get master from active document");

                    foreach (var name in this.Name)
                    {
                        var masters = this.Client.Master.GetMastersByNameFromDocument(name);
                        this.WriteObject(masters, true);
                    }
                }
            }
            else
            {
                // master name is not specified
                if (doc_specified)
                {
                    ((SMA.Cmdlet) this).WriteVerbose("Get all masters from specified document");
                    var masters = this.Client.Master.Get(this.Document);
                    this.WriteObject(masters, true);                    
                }
                else
                {
                    ((SMA.Cmdlet) this).WriteVerbose("Get all masters from active document");
                    var masters = this.Client.Master.Get();
                    this.WriteObject(masters, true);                   
                }
            }
        }
    }
}