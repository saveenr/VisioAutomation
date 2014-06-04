using System.Collections.Generic;
using System.Linq;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioApplication")]
    public class Get_VisioApplication : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            if (scriptingsession.VisioApplication  == null)
            {
                this.WriteVerboseEx("A Visio Application Instance is NOT Attached");
                this.WriteObject(null);
            }
            else
            {
                this.WriteVerboseEx("A Visio Application Instance is Attached");
                this.WriteObject(scriptingsession.VisioApplication);
            }
        }
    }
}