using System.Collections.Generic;
using System.Linq;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioApplication")]
    public class Get_VisioApplication : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            this.WriteObject(Globals.Application);
        }
    }
}