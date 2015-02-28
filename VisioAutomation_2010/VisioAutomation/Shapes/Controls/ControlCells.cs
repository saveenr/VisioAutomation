using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

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


        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return newpair(VA.ShapeSheet.SRCConstants.Controls_CanGlue, this.CanGlue.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Controls_Tip, this.Tip.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Controls_X, this.X.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Controls_Y, this.Y.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Controls_YCon, this.YBehavior.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Controls_XCon, this.XBehavior.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Controls_XDyn, this.XDynamics.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Controls_YDyn, this.YDynamics.Formula);
            }
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
            public CellColumn CanGlue { get; set; }
            public CellColumn Tip { get; set; }
            public CellColumn X { get; set; }
            public CellColumn Y { get; set; }
            public CellColumn YBehavior { get; set; }
            public CellColumn XBehavior { get; set; }
            public CellColumn XDynamics { get; set; }
            public CellColumn YDynamics { get; set; }

            public ControlCellQuery() 
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionControls);
                this.CanGlue = sec.AddCell(VA.ShapeSheet.SRCConstants.Controls_CanGlue);
                this.Tip = sec.AddCell(VA.ShapeSheet.SRCConstants.Controls_Tip);
                this.X = sec.AddCell(VA.ShapeSheet.SRCConstants.Controls_X);
                this.Y = sec.AddCell(VA.ShapeSheet.SRCConstants.Controls_Y);
                this.YBehavior = sec.AddCell(VA.ShapeSheet.SRCConstants.Controls_YCon);
                this.XBehavior = sec.AddCell(VA.ShapeSheet.SRCConstants.Controls_XCon);
                this.XDynamics = sec.AddCell(VA.ShapeSheet.SRCConstants.Controls_XDyn);
                this.YDynamics = sec.AddCell(VA.ShapeSheet.SRCConstants.Controls_YDyn);
            }

            public ControlCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
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