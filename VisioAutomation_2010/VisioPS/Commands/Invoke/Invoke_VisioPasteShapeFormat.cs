using System.Collections.Generic;
using VisioAutomation.Scripting;
using VisioAutomation.Shapes.Format;
using VA=VisioAutomation;
using SMA = System.Management.Automation;
using IVisio= Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioPasteFormat")]
    public class Invoke_VisioPasteShapeFormat : VisioPSCmdlet
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
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            FormatCategory category = 0x0;

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

            bool apply_formulas = false;
            scriptingsession.Format.Paste(this.Shapes,category,apply_formulas);
        }
    }
}