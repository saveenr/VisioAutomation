using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "SRCFormula")]
    public class Get_SRCFormula : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VisioAutomation.ShapeSheet.SRC[] SRC;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var formulas = scriptingsession.ShapeSheet.QueryFormulas(this.SRC);
            this.WriteObject(formulas);
        }
    }
}