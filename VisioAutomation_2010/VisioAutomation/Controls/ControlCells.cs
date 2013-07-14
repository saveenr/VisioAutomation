using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Controls
{
    public class ControlCells : VA.ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VA.ShapeSheet.CellData<int> CanGlue { get; set; }
        public VA.ShapeSheet.CellData<int> Tip { get; set; }
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> YBehavior { get; set; }
        public VA.ShapeSheet.CellData<int> XBehavior { get; set; }
        public VA.ShapeSheet.CellData<int> XDynamics { get; set; }
        public VA.ShapeSheet.CellData<int> YDynamics { get; set; }

        public override void ApplyFormulasForRow(ApplyFormula func, short row)
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

        public static IList<List<ControlCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells(page, shapeids, query, query.GetCells);
        }

        public static IList<ControlCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells(shape, query, query.GetCells);
        }

        private static ControlCellQuery _mCellQuery;
        private static ControlCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new ControlCellQuery();
            return _mCellQuery;
        }

        class ControlCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.CellQuery.Column CanGlue { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Tip { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column X { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Y { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column YBehavior { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column XBehavior { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column XDynamics { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column YDynamics { get; set; }

            public ControlCellQuery() 
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionControls);
                this.CanGlue = sec.AddColumn(VA.ShapeSheet.SRCConstants.Controls_CanGlue, "CanGlue");
                this.Tip = sec.AddColumn(VA.ShapeSheet.SRCConstants.Controls_Tip, "Tip");
                this.X = sec.AddColumn(VA.ShapeSheet.SRCConstants.Controls_X, "X");
                this.Y = sec.AddColumn(VA.ShapeSheet.SRCConstants.Controls_Y, "Y");
                this.YBehavior = sec.AddColumn(VA.ShapeSheet.SRCConstants.Controls_YCon, "YBehavior");
                this.XBehavior = sec.AddColumn(VA.ShapeSheet.SRCConstants.Controls_XCon, "XBehavior");
                this.XDynamics = sec.AddColumn(VA.ShapeSheet.SRCConstants.Controls_XDyn, "XDynamics");
                this.YDynamics = sec.AddColumn(VA.ShapeSheet.SRCConstants.Controls_YDyn, "YDynamics");
            }

            public ControlCells GetCells(VA.ShapeSheet.CellData<double>[] row)
            {
                var cells = new ControlCells();
                cells.CanGlue = row[CanGlue.Ordinal].ToInt();
                cells.Tip = row[Tip.Ordinal].ToInt();
                cells.X = row[X.Ordinal];
                cells.Y = row[Y.Ordinal];
                cells.YBehavior = row[YBehavior.Ordinal].ToInt();
                cells.XBehavior = row[XBehavior.Ordinal].ToInt();
                cells.XDynamics = row[XDynamics.Ordinal].ToInt();
                cells.YDynamics = row[YDynamics.Ordinal].ToInt();
                return cells;
            }
        }
    }
}