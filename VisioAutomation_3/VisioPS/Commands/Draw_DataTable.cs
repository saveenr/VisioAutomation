using System.Linq;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Draw", "DataTable")]
    public class Draw_DataTable : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = true)]
        public System.Data.DataTable DataTable { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public double CellWidth { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public double CellHeight { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public double CellSpacing { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var widths = Enumerable.Repeat<double>(CellWidth, DataTable.Columns.Count).ToList();
            var heights = Enumerable.Repeat<double>(CellHeight, DataTable.Rows.Count).ToList();
            var spacing = new VA.Drawing.Size(CellSpacing, CellSpacing);
            var shapes = scriptingsession.Draw.DrawDataTable(DataTable, widths, heights, spacing);
            this.WriteObject(shapes);
        }
    }
}