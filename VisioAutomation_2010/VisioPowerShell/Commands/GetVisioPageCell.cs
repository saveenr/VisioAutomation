using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using VisioPowerShell.Models;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioPageCell)]
    public class GetVisioPageCell : VisioCmdlet
    {
        [Parameter(Mandatory = false, Position = 0)]
        public string[] Cells { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Page Page { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter GetResults;

        [Parameter(Mandatory = false)]
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {
            var cellmap = VisioPowerShell.Models.PageCells.GetCellDictionary();

            this.Cells = cellmap.ExpandCellNames(this.Cells);

            var target_page = this.Page ?? this.Client.Page.Get();
            this.WriteVerbose("Valid Names: " + string.Join(",", cellmap.GetNames()));
            var query = cellmap.ToQuery(this.Cells);
            var surface = new VisioAutomation.SurfaceTarget(target_page);
            var target_shapeids = new[] { surface.Page.PageSheet.ID };
            var dt = DataTableHelpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);
        }
    }
}