using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioSRCResult")]
    public class Get_VisioSRCResult : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VisioAutomation.ShapeSheet.SRC [] SRC;

        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType= ResultType.Double;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            object results;
            if (this.ResultType == ResultType.Double)
            {
                results = scriptingsession.ShapeSheet.QueryResults<double>(this.SRC);
            }
            else if (this.ResultType == ResultType.Integer)
            {
                results = scriptingsession.ShapeSheet.QueryResults<int>(this.SRC);
            }
            else if (this.ResultType == ResultType.Boolean)
            {
                results = scriptingsession.ShapeSheet.QueryResults<bool>(this.SRC);
            }
            else if (this.ResultType == ResultType.String)
            {
                results = scriptingsession.ShapeSheet.QueryResults<string>(this.SRC);
            }
            else
            {
                results = scriptingsession.ShapeSheet.QueryResults<double>(this.SRC);
            }
            this.WriteObject(results);
        }
    }
}