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

            object results = null;
            if (this.ResultType == ResultType.Double)
            {
                results = scriptingsession.ShapeSheet.QueryResults<double>(this.Section, this.Cells);
            }
            else if (this.ResultType == ResultType.Integer)
            {
                results = scriptingsession.ShapeSheet.QueryResults<int>(this.Section, this.Cells);
            }
            else if (this.ResultType == ResultType.Boolean)
            {
                results = scriptingsession.ShapeSheet.QueryResults<bool>(this.Section, this.Cells);
            }
            else if (this.ResultType == ResultType.String)
            {
                results = scriptingsession.ShapeSheet.QueryResults<string>(this.Section, this.Cells);
            }
            else
            {
                results = scriptingsession.ShapeSheet.QueryResults<double>(this.Section, this.Cells);
            }
            this.WriteObject(results);

        }
    }
}