using IVisio=Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "SectionFormula")]
    public class Get_SectionFormula : VisioPS.VisioPSCmdlet
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

            dt = DTUtil.ToDataTable<string>(scriptingsession.ShapeSheet.QueryFormulas(this.Section, this.Cells));
            this.WriteObject(dt);
        }
    }
}