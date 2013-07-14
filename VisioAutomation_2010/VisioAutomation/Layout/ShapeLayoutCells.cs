using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.Layout
{
    public class ShapeLayoutCells : VA.ShapeSheet.CellGroups.CellGroup
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
        public VA.ShapeSheet.CellData<int> DisplayLevel { get; set; } // new in visio 2010
        public VA.ShapeSheet.CellData<int> Relationships { get; set; } // new in visio 2010

        public override void ApplyFormulas(ApplyFormula func)
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
            func(ShapeSheet.SRCConstants.DisplayLevel, this.DisplayLevel.Formula);
            func(ShapeSheet.SRCConstants.Relationships, this.Relationships.Formula);
        }



        public static IList<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells(page, shapeids, query, query.GetCells);
        }

        public static ShapeLayoutCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells(shape, query, query.GetCells);
        }

        private static ShapeLayoutCellQuery _mCellQuery;
        private static ShapeLayoutCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new ShapeLayoutCellQuery();
            return _mCellQuery;
        }

        class ShapeLayoutCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public Column ConFixedCode { get; set; }
            public Column ConLineJumpCode { get; set; }
            public Column ConLineJumpDirX { get; set; }
            public Column ConLineJumpDirY { get; set; }
            public Column ConLineJumpStyle { get; set; }
            public Column ConLineRouteExt { get; set; }
            public Column ShapeFixedCode { get; set; }
            public Column ShapePermeablePlace { get; set; }
            public Column ShapePermeableX { get; set; }
            public Column ShapePermeableY { get; set; }
            public Column ShapePlaceFlip { get; set; }
            public Column ShapePlaceStyle { get; set; }
            public Column ShapePlowCode { get; set; }
            public Column ShapeRouteStyle { get; set; }
            public Column ShapeSplit { get; set; }
            public Column ShapeSplittable { get; set; }
            public Column DisplayLevel { get; set; }
            public Column Relationships { get; set; }

            public ShapeLayoutCellQuery() :
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
                this.DisplayLevel= this.AddColumn(VA.ShapeSheet.SRCConstants.DisplayLevel, "DisplayLevel");
                this.Relationships = this.AddColumn(VA.ShapeSheet.SRCConstants.Relationships, "Relationships");
            }

            public ShapeLayoutCells GetCells(QueryResult<CellData<double>> data_for_shape)
            {
                var row = data_for_shape.Cells;
                var cells = new ShapeLayoutCells();
                cells.ConFixedCode = row[ConFixedCode.Ordinal].ToInt();
                cells.ConLineJumpCode = row[ConLineJumpCode.Ordinal].ToInt();
                cells.ConLineJumpDirX = row[ConLineJumpDirX.Ordinal].ToInt();
                cells.ConLineJumpDirY = row[ConLineJumpDirY.Ordinal].ToInt();
                cells.ConLineJumpStyle = row[ConLineJumpStyle.Ordinal].ToInt();
                cells.ConLineRouteExt = row[ConLineRouteExt.Ordinal].ToInt();
                cells.ShapeFixedCode = row[ShapeFixedCode.Ordinal].ToInt();
                cells.ShapePermeablePlace = row[ShapePermeablePlace.Ordinal].ToInt();
                cells.ShapePermeableX = row[ShapePermeableX.Ordinal].ToInt();
                cells.ShapePermeableY = row[ShapePermeableY.Ordinal].ToInt();
                cells.ShapePlaceFlip = row[ShapePlaceFlip.Ordinal].ToInt();
                cells.ShapePlaceStyle = row[ShapePlaceStyle.Ordinal].ToInt();
                cells.ShapePlowCode = row[ShapePlowCode.Ordinal].ToInt();
                cells.ShapeRouteStyle = row[ShapeRouteStyle.Ordinal].ToInt();
                cells.ShapeSplit = row[ShapeSplit.Ordinal].ToInt();
                cells.ShapeSplittable = row[ShapeSplittable.Ordinal].ToInt();
                cells.DisplayLevel = row[DisplayLevel.Ordinal].ToInt();
                cells.Relationships = row[Relationships.Ordinal].ToInt();
                return cells;
            }
        }

    }
}