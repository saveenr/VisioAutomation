using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes.Layout
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

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return newpair(ShapeSheet.SRCConstants.ConFixedCode, this.ConFixedCode.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ConLineJumpCode, this.ConLineJumpCode.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ConLineJumpDirX, this.ConLineJumpDirX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ConLineJumpDirY, this.ConLineJumpDirY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ConLineJumpStyle, this.ConLineJumpStyle.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ConLineRouteExt, this.ConLineRouteExt.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapeFixedCode, this.ShapeFixedCode.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapePermeablePlace, this.ShapePermeablePlace.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapePermeableX, this.ShapePermeableX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapePermeableY, this.ShapePermeableY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapePlaceFlip, this.ShapePlaceFlip.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapePlaceStyle, this.ShapePlaceStyle.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapePlowCode, this.ShapePlowCode.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapeRouteStyle, this.ShapeRouteStyle.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapeSplit, this.ShapeSplit.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapeSplittable, this.ShapeSplittable.Formula);
                yield return newpair(ShapeSheet.SRCConstants.DisplayLevel, this.DisplayLevel.Formula);
                yield return newpair(ShapeSheet.SRCConstants.Relationships, this.Relationships.Formula);
            }
        }


        public static IList<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<ShapeLayoutCells, double>(page, shapeids, query, query.GetCells);
        }

        public static ShapeLayoutCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<ShapeLayoutCells, double>(shape, query, query.GetCells);
        }

        private static ShapeLayoutCellQuery _mCellQuery;
        private static ShapeLayoutCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new ShapeLayoutCellQuery();
            return _mCellQuery;
        }

        class ShapeLayoutCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public CellColumn ConFixedCode { get; set; }
            public CellColumn ConLineJumpCode { get; set; }
            public CellColumn ConLineJumpDirX { get; set; }
            public CellColumn ConLineJumpDirY { get; set; }
            public CellColumn ConLineJumpStyle { get; set; }
            public CellColumn ConLineRouteExt { get; set; }
            public CellColumn ShapeFixedCode { get; set; }
            public CellColumn ShapePermeablePlace { get; set; }
            public CellColumn ShapePermeableX { get; set; }
            public CellColumn ShapePermeableY { get; set; }
            public CellColumn ShapePlaceFlip { get; set; }
            public CellColumn ShapePlaceStyle { get; set; }
            public CellColumn ShapePlowCode { get; set; }
            public CellColumn ShapeRouteStyle { get; set; }
            public CellColumn ShapeSplit { get; set; }
            public CellColumn ShapeSplittable { get; set; }
            public CellColumn DisplayLevel { get; set; }
            public CellColumn Relationships { get; set; }

            public ShapeLayoutCellQuery() :
                base()
            {
                this.ConFixedCode = this.AddCell(VA.ShapeSheet.SRCConstants.ConFixedCode, "ConFixedCode");
                this.ConLineJumpCode = this.AddCell(VA.ShapeSheet.SRCConstants.ConLineJumpCode, "ConLineJumpCode");
                this.ConLineJumpDirX = this.AddCell(VA.ShapeSheet.SRCConstants.ConLineJumpDirX, "ConLineJumpDirX");
                this.ConLineJumpDirY = this.AddCell(VA.ShapeSheet.SRCConstants.ConLineJumpDirY, "ConLineJumpDirY");
                this.ConLineJumpStyle = this.AddCell(VA.ShapeSheet.SRCConstants.ConLineJumpStyle, "ConLineJumpStyle");
                this.ConLineRouteExt = this.AddCell(VA.ShapeSheet.SRCConstants.ConLineRouteExt, "ConLineRouteExt");
                this.ShapeFixedCode = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeFixedCode, "ShapeFixedCode");
                this.ShapePermeablePlace = this.AddCell(VA.ShapeSheet.SRCConstants.ShapePermeablePlace, "ShapePermeablePlace");
                this.ShapePermeableX = this.AddCell(VA.ShapeSheet.SRCConstants.ShapePermeableX, "ShapePermeableX");
                this.ShapePermeableY = this.AddCell(VA.ShapeSheet.SRCConstants.ShapePermeableY, "ShapePermeableY");
                this.ShapePlaceFlip = this.AddCell(VA.ShapeSheet.SRCConstants.ShapePlaceFlip, "ShapePlaceFlip");
                this.ShapePlaceStyle = this.AddCell(VA.ShapeSheet.SRCConstants.ShapePlaceStyle, "ShapePlaceStyle");
                this.ShapePlowCode = this.AddCell(VA.ShapeSheet.SRCConstants.ShapePlowCode, "ShapePlowCode");
                this.ShapeRouteStyle = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeRouteStyle, "ShapeRouteStyle");
                this.ShapeSplit = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeSplit, "ShapeSplit");
                this.ShapeSplittable = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeSplittable, "ShapeSplittable");
                this.DisplayLevel= this.AddCell(VA.ShapeSheet.SRCConstants.DisplayLevel, "DisplayLevel");
                this.Relationships = this.AddCell(VA.ShapeSheet.SRCConstants.Relationships, "Relationships");
            }

            public ShapeLayoutCells GetCells(VA.ShapeSheet.CellData<double>[] row)
            {
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