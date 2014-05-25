using System.Collections.Generic;
using System.Linq;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioSession")]
    public class Get_VisioSession : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var ss = this.ScriptingSession;
            this.WriteObject(ss);
        }
    }
}