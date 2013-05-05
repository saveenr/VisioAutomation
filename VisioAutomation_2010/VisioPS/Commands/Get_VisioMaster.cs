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

        [SMA.Parameter(Position = 1, Mandatory = false, ParameterSetName = "StencilName")]
        public IVisio.Document Stencil;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            bool isprovided_Master = this.Master !=null;
            bool isprovided_Stencil = this.Stencil !=null;

            this.WriteVerboseEx("Master name provided: {0}", isprovided_Master);
            this.WriteVerboseEx("Stencil document provided: {0}", isprovided_Stencil);

            if (!isprovided_Master && isprovided_Stencil)
            {
                this.WriteVerbose("Retrieve all the masters in the specified stencil doc");
                var masters = scriptingsession.Master.Get(this.Stencil);
                this.WriteObject(masters,true);
            }
            else if (isprovided_Master && !isprovided_Stencil)
            {
                this.WriteVerbose("Retrieve a specific master in the currently active document");
                var master = scriptingsession.Master.Get(this.Master);
                this.WriteObject(master);
            }
            else if (!isprovided_Master && !isprovided_Stencil)
            {
                this.WriteVerbose("Retrieve all the masters in the currently active document");
                var masters = scriptingsession.Master.Get();
                this.WriteObject(masters, true);
                return;
            }
            else if (isprovided_Master && isprovided_Stencil)
            {
                this.WriteVerbose("Retrieve a specific master in the specified stencil document");
                var master = scriptingsession.Master.Get(this.Master, this.Stencil);
                this.WriteObject(master);
            }           
        }
    }
}