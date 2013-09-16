using System.Collections.Generic;
using VisioAutomation.Scripting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioCopyShapeFormat")]
    public class Invoke_VisioCopyShapeFormat : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Fill { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Line { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Shadow { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Text { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public Microsoft.Office.Interop.Visio.Shape Shape;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            FormatCategory category = 0x00;

            if (Fill)
            {
                category |= FormatCategory.Fill;
            }
            if (Line)
            {
                category |= FormatCategory.Line;
            }
            if (Shadow)
            {
                category |= FormatCategory.Shadow;
            }
            if (Text)
            {
                category |= FormatCategory.Character;
            }
            
            scriptingsession.Format.Copy(this.Shape, category);
        }
    }
}