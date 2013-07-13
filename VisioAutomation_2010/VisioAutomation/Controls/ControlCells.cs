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

        private static ControlCells get_cells_from_row(ControlQuery query, VA.ShapeSheet.Data.Table<VA.ShapeSheet.CellData<double>> table, int row)
        {
            var cells = new ControlCells();
            cells.CanGlue = table[row,query.CanGlue].ToInt();
            cells.Tip = table[row,query.Tip].ToInt();
            cells.X = table[row,query.X];
            cells.Y = table[row,query.Y];
            cells.YBehavior = table[row,query.YBehavior].ToInt();
            cells.XBehavior = table[row,query.XBehavior].ToInt();
            cells.XDynamics = table[row,query.XDynamics].ToInt();
            cells.YDynamics = table[row,query.YDynamics].ToInt();
            return cells;
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

        private static ControlQuery m_query;
        private static ControlQuery get_query()
        {
            m_query = m_query ?? new ControlQuery();
            return m_query;
        }

        class ControlQuery : VA.ShapeSheet.Query.QueryEx
        {
            public int CanGlue { get; set; }
            public int Tip { get; set; }
            public int X { get; set; }
            public int Y { get; set; }
            public int YBehavior { get; set; }
            public int XBehavior { get; set; }
            public int XDynamics { get; set; }
            public int YDynamics { get; set; }

            public ControlQuery() 
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
                cells.CanGlue = row[CanGlue].ToInt();
                cells.Tip = row[Tip].ToInt();
                cells.X = row[X];
                cells.Y = row[Y];
                cells.YBehavior = row[YBehavior].ToInt();
                cells.XBehavior = row[XBehavior].ToInt();
                cells.XDynamics = row[XDynamics].ToInt();
                cells.YDynamics = row[YDynamics].ToInt();
                return cells;
            }
        }
    }
}