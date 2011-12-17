using IVisio=Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "SectionResult")]
    public class Get_SectionResult : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.VisSectionIndices Section;

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public IVisio.VisCellIndices[] Cells;
        
        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType = ResultType.Double;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            System.Data.DataTable dt = null;
            if (this.ResultType == ResultType.Double)
            {
                dt = DTUtil.ToDataTable<double>(scriptingsession.ShapeSheet.QueryResults<double>(this.Section, this.Cells));
            }
            else if (this.ResultType == ResultType.Integer)
            {
                dt = DTUtil.ToDataTable<int>(scriptingsession.ShapeSheet.QueryResults<int>(this.Section, this.Cells));
            }
            else if (this.ResultType == ResultType.Boolean)
            {
                dt = DTUtil.ToDataTable<bool>(scriptingsession.ShapeSheet.QueryResults<bool>(this.Section, this.Cells));
            }
            else if (this.ResultType == ResultType.String)
            {
                dt = DTUtil.ToDataTable<string>(scriptingsession.ShapeSheet.QueryResults<string>(this.Section, this.Cells));
            }
            else
            {
                throw new System.ArgumentOutOfRangeException("ResultType");
            }
            this.WriteObject(dt);

        }
    }
}