using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Shapes.Controls
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


        public override IEnumerable<SRCValuePair> EnumPairs()
        {
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.Controls_CanGlue, this.CanGlue.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.Controls_Tip, this.Tip.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.Controls_X, this.X.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.Controls_Y, this.Y.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.Controls_YCon, this.YBehavior.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.Controls_XCon, this.XBehavior.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.Controls_XDyn, this.XDynamics.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.Controls_YDyn, this.YDynamics.Formula);
        }

        public static IList<List<ControlCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells<ControlCells,double>(page, shapeids, query, query.GetCells);
        }

        public static IList<ControlCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells<ControlCells,double>(shape, query, query.GetCells);
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
                var sec = this.Sections.Add(IVisio.VisSectionIndices.visSectionControls);
                this.CanGlue = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Controls_CanGlue, "CanGlue");
                this.Tip = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Controls_Tip, "Tip");
                this.X = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Controls_X, "X");
                this.Y = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Controls_Y, "Y");
                this.YBehavior = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Controls_YCon, "YBehavior");
                this.XBehavior = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Controls_XCon, "XBehavior");
                this.XDynamics = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Controls_XDyn, "XDynamics");
                this.YDynamics = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Controls_YDyn, "YDynamics");
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