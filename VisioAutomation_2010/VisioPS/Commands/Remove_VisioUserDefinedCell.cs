using System.Collections.Generic;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioUserDefinedCell")]
    public class Remove_VisioUserDefinedCell : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IList<Microsoft.Office.Interop.Visio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            scriptingsession.UserDefinedCell.Delete(this.Shapes,Name);
        }
    }
}