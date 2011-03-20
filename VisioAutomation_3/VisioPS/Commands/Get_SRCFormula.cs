using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "SRCFormula")]
    public class Get_SRCFormula : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VisioAutomation.ShapeSheet.SRC[] SRC;

        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType = ResultType.Double;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            System.Data.DataTable dt = null;
            dt = DTUtil.ToDataTable<string>(scriptingsession.ShapeSheet.QueryFormulas(this.SRC));
            this.WriteObject(dt);
        }
    }
}