using System.Collections.Generic;
using VA=VisioAutomation;
using VisioPS.Extensions;
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
        public IList<IVisio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            VA.Format.FormatCategory category = 0x0;

            if (Fill)
            {
                category |= VA.Format.FormatCategory.Fill;
            }
            if (Line)
            {
                category |= VA.Format.FormatCategory.Line;
            }
            if (Shadow)
            {
                category |= VA.Format.FormatCategory.Shadow;
            }
            if (Text)
            {
                category |= VA.Format.FormatCategory.Character;
            }

            bool apply_formulas = false;
            scriptingsession.Format.PasteFormat(this.Shapes,category,apply_formulas);
        }
    }
}