using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, "VisioPageCell")]
    public class Get_VisioPageCell : VisioCmdlet
    {
        [Parameter(Mandatory = false, Position = 0)]
        public string[] Cells { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Page Page { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter GetResults;

        [Parameter(Mandatory = false)]
        public Model.ResultType ResultType = Model.ResultType.String;

        protected override void ProcessRecord()
        {
            var cellmap = CellSRCDictionary.GetCellMapForPages();
            if (this.Cells == null || this.Cells.Length < 1 || this.Cells.Contains("*"))
            {
                this.Cells = cellmap.GetNames().ToArray();
            }
            Get_VisioPageCell.EnsureEnoughCellNames(this.Cells);
            var target_page = this.Page ?? this.client.Page.Get();
            this.WriteVerbose("Valid Names: " + string.Join(",", cellmap.GetNames()));
            var query = cellmap.CreateQueryFromCellNames(this.Cells);
            var surface = new VA.ShapeSheet.ShapeSheetSurface(target_page);
            var target_shapeids = new[] { surface.Target.Page.PageSheet.ID };
            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);
        }

        public static void EnsureEnoughCellNames(IList<string> Cells)
        {
            if (Cells == null)
            {
                throw new ArgumentNullException(nameof(Cells));
            }

            if (Cells.Count< 1)
            {
                string msg = "Must provide at least one cell name";
                throw new ArgumentException(msg,nameof(Cells));
            }
        }
    }
}