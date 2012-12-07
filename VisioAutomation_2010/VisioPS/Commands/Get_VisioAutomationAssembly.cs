using System.Collections.Generic;
using System.Linq;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioAutomationAssembly")]
    public class VisioAutomationAssembly : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var type = typeof (VA.ShapeSheet.SRC);
            var asm = type.Assembly;
            this.WriteObject(asm);
        }
    }
}