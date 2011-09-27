using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.Layout
{
    public partial class ShapeLayoutCells : VA.ShapeSheet.CellGroup
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

        protected override void _Apply(VA.ShapeSheet.CellGroup.ApplyFormula func)
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

        private static ShapeLayoutCells get_cells_from_row(ShapeLayoutQuery query, VA.ShapeSheet.Data.QueryDataRow<double> row)
        {
            var cells = new ShapeLayoutCells();
            cells.ConFixedCode = row[query.ConFixedCode].ToInt();
            cells.ConLineJumpCode = row[query.ConLineJumpCode].ToInt();
            cells.ConLineJumpDirX = row[query.ConLineJumpDirX].ToInt();
            cells.ConLineJumpDirY = row[query.ConLineJumpDirY].ToInt();
            cells.ConLineJumpStyle = row[query.ConLineJumpStyle].ToInt();
            cells.ConLineRouteExt = row[query.ConLineRouteExt].ToInt();
            cells.ShapeFixedCode = row[query.ShapeFixedCode].ToInt();
            cells.ShapePermeablePlace = row[query.ShapePermeablePlace].ToInt();
            cells.ShapePermeableX = row[query.ShapePermeableX].ToInt();
            cells.ShapePermeableY = row[query.ShapePermeableY].ToInt();
            cells.ShapePlaceFlip = row[query.ShapePlaceFlip].ToInt();
            cells.ShapePlaceStyle = row[query.ShapePlaceStyle].ToInt();
            cells.ShapePlowCode = row[query.ShapePlowCode].ToInt();
            cells.ShapeRouteStyle = row[query.ShapeRouteStyle].ToInt();
            cells.ShapeSplit = row[query.ShapeSplit].ToInt();
            cells.ShapeSplittable = row[query.ShapeSplittable].ToInt();
            return cells;
        }

        internal static IList<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new ShapeLayoutQuery();
            return VA.ShapeSheet.CellGroup._GetObjectsFromRows(page, shapeids, query, get_cells_from_row);
        }

        internal static ShapeLayoutCells GetCells(IVisio.Shape shape)
        {
            var query = new ShapeLayoutQuery();
            return VA.ShapeSheet.CellGroup._GetObjectFromSingleRow(shape, query, get_cells_from_row);
        }

        class ShapeLayoutQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.CellQueryColumn ConFixedCode { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ConLineJumpCode { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ConLineJumpDirX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ConLineJumpDirY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ConLineJumpStyle { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ConLineRouteExt { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapeFixedCode { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapePermeablePlace { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapePermeableX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapePermeableY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapePlaceFlip { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapePlaceStyle { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapePlowCode { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapeRouteStyle { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapeSplit { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapeSplittable { get; set; }

            public ShapeLayoutQuery() :
                base()
            {
                this.ConFixedCode = this.AddColumn(VA.ShapeSheet.SRCConstants.ConFixedCode, "ConFixedCode");
                this.ConLineJumpCode = this.AddColumn(VA.ShapeSheet.SRCConstants.ConLineJumpCode, "ConLineJumpCode");
                this.ConLineJumpDirX = this.AddColumn(VA.ShapeSheet.SRCConstants.ConLineJumpDirX, "ConLineJumpDirX");
                this.ConLineJumpDirY = this.AddColumn(VA.ShapeSheet.SRCConstants.ConLineJumpDirY, "ConLineJumpDirY");
                this.ConLineJumpStyle = this.AddColumn(VA.ShapeSheet.SRCConstants.ConLineJumpStyle, "ConLineJumpStyle");
                this.ConLineRouteExt = this.AddColumn(VA.ShapeSheet.SRCConstants.ConLineRouteExt, "ConLineRouteExt");
                this.ShapeFixedCode = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeFixedCode, "ShapeFixedCode");
                this.ShapePermeablePlace = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePermeablePlace, "ShapePermeablePlace");
                this.ShapePermeableX = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePermeableX, "ShapePermeableX");
                this.ShapePermeableY = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePermeableY, "ShapePermeableY");
                this.ShapePlaceFlip = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePlaceFlip, "ShapePlaceFlip");
                this.ShapePlaceStyle = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePlaceStyle, "ShapePlaceStyle");
                this.ShapePlowCode = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePlowCode, "ShapePlowCode");
                this.ShapeRouteStyle = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeRouteStyle, "ShapeRouteStyle");
                this.ShapeSplit = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeSplit, "ShapeSplit");
                this.ShapeSplittable = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeSplittable, "ShapeSplittable");
            }
        }

    }
}