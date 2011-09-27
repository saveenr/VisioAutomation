using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Controls
{
    public class ControlCells : VA.ShapeSheet.CellGroups.CellGroupForSection
    {
        public VA.ShapeSheet.CellData<int> CanGlue { get; set; }
        public VA.ShapeSheet.CellData<int> Tip { get; set; }
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> YBehavior { get; set; }
        public VA.ShapeSheet.CellData<int> XBehavior { get; set; }
        public VA.ShapeSheet.CellData<int> XDynamics { get; set; }
        public VA.ShapeSheet.CellData<int> YDynamics { get; set; }

        protected override void _Apply(ApplyFormula func, short row)
        {
            func(VA.ShapeSheet.SRCConstants.Controls_CanGlue.ForRow(row), this.CanGlue.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_Tip.ForRow(row), this.Tip.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_X.ForRow(row), this.X.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_Y.ForRow(row), this.Y.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_YCon.ForRow(row), this.YBehavior.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_XCon.ForRow(row), this.XBehavior.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_XDyn.ForRow(row), this.XDynamics.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_YDyn.ForRow(row), this.YDynamics.Formula);
        }

        private static ControlCells get_cells_from_row(ControlQuery query, VA.ShapeSheet.Data.QueryDataRow<double> row)
        {
            var cells = new ControlCells();
            cells.CanGlue = row[query.CanGlue].ToInt();
            cells.Tip = row[query.Tip].ToInt();
            cells.X = row[query.X];
            cells.Y = row[query.Y];
            cells.YBehavior = row[query.YBehavior].ToInt();
            cells.XBehavior = row[query.XBehavior].ToInt();
            cells.XDynamics = row[query.XDynamics].ToInt();
            cells.YDynamics = row[query.YDynamics].ToInt();
            return cells;
        }

        internal static IList<List<ControlCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new ControlQuery();
            return VA.ShapeSheet.CellGroups.CellGroupForSection._GetObjectsFromRowsGrouped(page, shapeids, query, get_cells_from_row);
        }

        internal static IList<ControlCells> GetCells(IVisio.Shape shape)
        {
            var query = new ControlQuery();
            return VA.ShapeSheet.CellGroups.CellGroupForSection._GetObjectsFromRows(shape, query, get_cells_from_row);
        }

        class ControlQuery : VA.ShapeSheet.Query.SectionQuery
        {
            public VA.ShapeSheet.Query.SectionQueryColumn CanGlue { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn Tip { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn X { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn Y { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn YBehavior { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn XBehavior { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn XDynamics { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn YDynamics { get; set; }

            public ControlQuery() :
                base(IVisio.VisSectionIndices.visSectionControls)
            {
                this.CanGlue = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_CanGlue.Cell, "CanGlue");
                this.Tip = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_Tip.Cell, "Tip");
                this.X = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_X.Cell, "X");
                this.Y = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_Y.Cell, "Y");
                this.YBehavior = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_YCon.Cell, "YBehavior");
                this.XBehavior = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_XCon.Cell, "XBehavior");
                this.XDynamics = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_XDyn.Cell, "XDynamics");
                this.YDynamics = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_YDyn.Cell, "YDynamics");
            }
        }
    }
}