using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, "VisioMaster")]
    public class Get_VisioMaster : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = false)]
        public string Name;

        [Parameter(Position = 1, Mandatory = false)]
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
                    ((Cmdlet) this).WriteVerbose("Get master from specified document");
                    var masters = this.client.Master.GetMastersByName(this.Name, this.Document);
                    this.WriteObject(masters,true);
                }
                else
                {
                    ((Cmdlet) this).WriteVerbose("Get master from active document");
                    var masters = this.client.Master.GetMastersByName(this.Name);
                    this.WriteObject(masters,true);
                }
            }
            else
            {
                // master name is not specified
                if (doc_specified)
                {
                    ((Cmdlet) this).WriteVerbose("Get all masters from specified document");
                    var masters = this.client.Master.Get(this.Document);
                    this.WriteObject(masters, false);                    
                }
                else
                {
                    ((Cmdlet) this).WriteVerbose("Get all masters from active document");
                    var masters = this.client.Master.Get();
                    this.WriteObject(masters, false);                   
                }
            }
        }
    }
}