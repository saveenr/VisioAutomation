using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.Layout
{
    public partial class ShapeLayoutCells : VA.ShapeSheet.CellDataGroup
    {
        public VA.ShapeSheet.CellData<int> ConFixedCode { get; set; }
        public VA.ShapeSheet.CellData<int> ConLineJumpCode { get; set; }
        public VA.ShapeSheet.CellData<int> ConLineJumpDirX { get; set; }
        public VA.ShapeSheet.CellData<int> ConLineJumpDirY { get; set; }
        public VA.ShapeSheet.CellData<int> ConLineJumpStyle { get; set; }
        public VA.ShapeSheet.CellData<int> ConLineRouteExt { get; set; }
        public VA.ShapeSheet.CellData<int> ShapeFixedCode { get; set; }
        public VA.ShapeSheet.CellData<int> ShapePermeablePlace { get; set; }
        public VA.ShapeSheet.CellData<int> ShapePermeableX { get; set; }
        public VA.ShapeSheet.CellData<int> ShapePermeableY { get; set; }
        public VA.ShapeSheet.CellData<int> ShapePlaceFlip { get; set; }
        public VA.ShapeSheet.CellData<int> ShapePlaceStyle { get; set; }
        public VA.ShapeSheet.CellData<int> ShapePlowCode { get; set; }
        public VA.ShapeSheet.CellData<int> ShapeRouteStyle { get; set; }
        public VA.ShapeSheet.CellData<int> ShapeSplit { get; set; }
        public VA.ShapeSheet.CellData<int> ShapeSplittable { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellDataGroup.ApplyFormula func)
        {
            func(ShapeSheet.SRCConstants.ConFixedCode, this.ConFixedCode.Formula);
            func(ShapeSheet.SRCConstants.ConLineJumpCode, this.ConLineJumpCode.Formula);
            func(ShapeSheet.SRCConstants.ConLineJumpDirX, this.ConLineJumpDirX.Formula);
            func(ShapeSheet.SRCConstants.ConLineJumpDirY, this.ConLineJumpDirY.Formula);
            func(ShapeSheet.SRCConstants.ConLineJumpStyle, this.ConLineJumpStyle.Formula);
            func(ShapeSheet.SRCConstants.ConLineRouteExt, this.ConLineRouteExt.Formula);
            func(ShapeSheet.SRCConstants.ShapeFixedCode, this.ShapeFixedCode.Formula);
            func(ShapeSheet.SRCConstants.ShapePermeablePlace, this.ShapePermeablePlace.Formula);
            func(ShapeSheet.SRCConstants.ShapePermeableX, this.ShapePermeableX.Formula);
            func(ShapeSheet.SRCConstants.ShapePermeableY, this.ShapePermeableY.Formula);
            func(ShapeSheet.SRCConstants.ShapePlaceFlip, this.ShapePlaceFlip.Formula);
            func(ShapeSheet.SRCConstants.ShapePlaceStyle, this.ShapePlaceStyle.Formula);
            func(ShapeSheet.SRCConstants.ShapePlowCode, this.ShapePlowCode.Formula);
            func(ShapeSheet.SRCConstants.ShapeRouteStyle, this.ShapeRouteStyle.Formula);
            func(ShapeSheet.SRCConstants.ShapeSplit, this.ShapeSplit.Formula);
            func(ShapeSheet.SRCConstants.ShapeSplittable, this.ShapeSplittable.Formula);
        }

        private static ShapeLayoutCells get_cells_from_row(ShapeLayoutQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new ShapeLayoutCells();
            cells.ConFixedCode = qds.GetItem(row, query.ConFixedCode).Cast(v => (int)v);
            cells.ConLineJumpCode = qds.GetItem(row, query.ConLineJumpCode).Cast(v => (int)v);
            cells.ConLineJumpDirX = qds.GetItem(row, query.ConLineJumpDirX).Cast(v => (int)v);
            cells.ConLineJumpDirY = qds.GetItem(row, query.ConLineJumpDirY).Cast(v => (int)v);
            cells.ConLineJumpStyle = qds.GetItem(row, query.ConLineJumpStyle).Cast(v => (int)v);
            cells.ConLineRouteExt = qds.GetItem(row, query.ConLineRouteExt).Cast(v => (int)v);
            cells.ShapeFixedCode = qds.GetItem(row, query.ShapeFixedCode).Cast(v => (int)v);
            cells.ShapePermeablePlace = qds.GetItem(row, query.ShapePermeablePlace).Cast(v => (int)v);
            cells.ShapePermeableX = qds.GetItem(row, query.ShapePermeableX).Cast(v => (int)v);
            cells.ShapePermeableY = qds.GetItem(row, query.ShapePermeableY).Cast(v => (int)v);
            cells.ShapePlaceFlip = qds.GetItem(row, query.ShapePlaceFlip).Cast(v => (int)v);
            cells.ShapePlaceStyle = qds.GetItem(row, query.ShapePlaceStyle).Cast(v => (int)v);
            cells.ShapePlowCode = qds.GetItem(row, query.ShapePlowCode).Cast(v => (int)v);
            cells.ShapeRouteStyle = qds.GetItem(row, query.ShapeRouteStyle).Cast(v => (int)v);
            cells.ShapeSplit = qds.GetItem(row, query.ShapeSplit).Cast(v => (int)v);
            cells.ShapeSplittable = qds.GetItem(row, query.ShapeSplittable).Cast(v => (int)v);
            return cells;
        }

        internal static IList<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new ShapeLayoutQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        internal static ShapeLayoutCells GetCells(IVisio.Shape shape)
        {
            var query = new ShapeLayoutQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(shape, query, get_cells_from_row);
        }
    }
}