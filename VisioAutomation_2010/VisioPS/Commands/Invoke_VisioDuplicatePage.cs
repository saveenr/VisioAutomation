using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioDuplicatePage")]
    public class Invoke_VisioDuplicatePage : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name = null;

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public IVisio.Document ToDocument=null;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.ToDocument == null)
            {
                scriptingsession.Page.Duplicate(this.Name);
            }
            else
            {
                scriptingsession.Page.Duplicate(this.Name, this.ToDocument);
            }
        }
    }
}