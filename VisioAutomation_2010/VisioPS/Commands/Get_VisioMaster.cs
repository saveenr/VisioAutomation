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

            if (!isprovided(Master) && isprovided(Stencil))
            {
                // Master name is not provided
                // Stencil name is provided
                // So retrieve all the masters in that stencil doc
                var masters = scriptingsession.Master.Get(this.Stencil);
                this.WriteObject(masters);
            }
            else if (isprovided(Master) && !isprovided(Stencil))
            {
                // Master was given
                // Stencil was not given
                // return that master in the current doc
                var master = scriptingsession.Master.Get(this.Master);
                this.WriteObject(master);
            }
            else if (!isprovided(Master) && !isprovided(Stencil))
            {
                // Neither was given, return all the masters in the active doc
                var masters = scriptingsession.Master.Get();
                this.WriteObject(masters);
                return;
            }
            else if (isprovided(Master) && isprovided(Stencil))
            {
                // Master & Stencil were given, retrive the master in that stencil
                var master = scriptingsession.Master.Get(this.Master, this.Stencil);
                this.WriteObject(master);
            }           
        }

        private bool isprovided<T>(T s) where T : class 
        {
            return s != null;
        }
    }
}