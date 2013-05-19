using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioDuplicateShape")]
    public class Invoke_VisioDuplicateShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Selection.Duplicate(this.Shapes);
        }
    }
}