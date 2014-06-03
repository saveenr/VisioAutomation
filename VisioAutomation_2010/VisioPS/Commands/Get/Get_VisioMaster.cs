using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioMaster")]
    public class Get_VisioMaster : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Master;

        [SMA.Parameter(Position = 1, Mandatory = false, ParameterSetName = "StencilDoc")]
        public IVisio.Document Stencil;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            bool master_specified = this.Master !=null;
            bool stencil_specified = this.Stencil !=null;

            if (master_specified)
            {
                if (stencil_specified)
                {
                    this.WriteVerbose("Get master from specified document");
                    var master = scriptingsession.Master.Get(this.Master, this.Stencil);
                    this.WriteObject(master);
                }
                else
                {
                    this.WriteVerbose("Get master from active document");
                    var master = scriptingsession.Master.Get(this.Master);
                    this.WriteObject(master);
                }
            }
            else
            {
                if (stencil_specified)
                {
                    this.WriteVerbose("Get all masters from specified document");
                    var masters = scriptingsession.Master.Get(this.Stencil);
                    this.WriteObject(masters, false);                    
                }
                else
                {
                    this.WriteVerbose("Get all masters from active document");
                    var masters = scriptingsession.Master.Get();
                    this.WriteObject(masters, false);                   
                }
            }
        }
    }
}