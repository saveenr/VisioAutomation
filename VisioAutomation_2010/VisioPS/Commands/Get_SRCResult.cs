using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "SRCResult")]
    public class Get_SRCResult : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VisioAutomation.ShapeSheet.SRC [] SRC;

        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType= ResultType.Double;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            System.Data.DataTable dt = null;
            if (this.ResultType == ResultType.Double)
            {
                dt = DataUtil.ToDataTable<double>(scriptingsession.ShapeSheet.QueryResults<double>(this.SRC));
            }
            else if (this.ResultType == ResultType.Integer)
            {
                dt = DataUtil.ToDataTable<int>(scriptingsession.ShapeSheet.QueryResults<int>(this.SRC));
            }
            else if (this.ResultType == ResultType.Boolean)
            {
                dt = DataUtil.ToDataTable<bool>(scriptingsession.ShapeSheet.QueryResults<bool>(this.SRC));
            }
            else if (this.ResultType == ResultType.String)
            {
                dt = DataUtil.ToDataTable<string>(scriptingsession.ShapeSheet.QueryResults<string>(this.SRC));
            }
            else
            {
                throw new System.ArgumentOutOfRangeException("ResultType");
            }
            this.WriteObject(dt);
        }
    }
}