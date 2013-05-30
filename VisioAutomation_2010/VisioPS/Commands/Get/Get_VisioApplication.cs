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
            if (this.AttachedVisioApplication == null)
            {
                this.WriteVerboseEx("A Visio Application Instance is Attached");
            }
            else
            {
                this.WriteVerboseEx("A Visio Application Instance is not Attached");                
            }
            this.WriteObject(AttachedVisioApplication);
        }
    }
}