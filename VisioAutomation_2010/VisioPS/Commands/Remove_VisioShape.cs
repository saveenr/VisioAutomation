using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioShape")]
    public class Remove_VisioShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IList<IVisio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            if (this.Shapes == null)
            {
                scriptingsession.Selection.Delete();                
            }
            else
            {
                foreach (var shape in this.Shapes)
                {
                    shape.Delete();
                }
            }
        }
    }
}