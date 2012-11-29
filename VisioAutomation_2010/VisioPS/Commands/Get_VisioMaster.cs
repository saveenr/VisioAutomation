using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "VisioMaster")]
    public class Get_VisioMaster : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Master;

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string Stencil;


        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            
            if (Master == null && Stencil != null)
            {
                // Master name is not provided
                // Stencil name is provided
                // So retrieve all the masters in that stenil doc
                var doc = scriptingsession.Document.Get(Stencil);
                var masters = scriptingsession.Master.Get(doc);
                this.WriteObject(masters);
            }
            else if (Master != null && Stencil == null)
            {
                // Master was given
                // Stencil was not given
                // return that master in the current doc
                var master = scriptingsession.Master.Get(this.Master);
                this.WriteObject(master);
            }
            else if (Master == null && Stencil == null)
            {
                // Neither was given, return all the masters in the active doc
                var masters = scriptingsession.Master.Get();
                this.WriteObject(masters);
                return;
            }
            else if (Master != null && Stencil != null)
            {
                // Master & Stencil were given, retrive the master in that stencil
                var master = scriptingsession.Master.Get(this.Master, this.Stencil);
                this.WriteObject(master);
            }           
        }
    }
}