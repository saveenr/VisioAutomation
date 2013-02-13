using System.Collections.Generic;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioText")]
    public class Set_VisioText : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string[] Text { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IList<Microsoft.Office.Interop.Visio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Text.SetText(this.Shapes,Text);
        }
    }
}
