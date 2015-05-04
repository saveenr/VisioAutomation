using System.Collections.Generic;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes.Layout
{
    public class ShapeLayoutCells : ShapeSheet.CellGroups.CellGroup
    {
        public ShapeSheet.CellData<int> ConFixedCode { get; set; }
        public ShapeSheet.CellData<int> ConLineJumpCode { get; set; }
        public ShapeSheet.CellData<int> ConLineJumpDirX { get; set; }
        public ShapeSheet.CellData<int> ConLineJumpDirY { get; set; }
        public ShapeSheet.CellData<int> ConLineJumpStyle { get; set; }
        public ShapeSheet.CellData<int> ConLineRouteExt { get; set; }
        public ShapeSheet.CellData<int> ShapeFixedCode { get; set; }
        public ShapeSheet.CellData<int> ShapePermeablePlace { get; set; }
        public ShapeSheet.CellData<int> ShapePermeableX { get; set; }
        public ShapeSheet.CellData<int> ShapePermeableY { get; set; }
        public ShapeSheet.CellData<int> ShapePlaceFlip { get; set; }
        public ShapeSheet.CellData<int> ShapePlaceStyle { get; set; }
        public ShapeSheet.CellData<int> ShapePlowCode { get; set; }
        public ShapeSheet.CellData<int> ShapeRouteStyle { get; set; }
        public ShapeSheet.CellData<int> ShapeSplit { get; set; }
        public ShapeSheet.CellData<int> ShapeSplittable { get; set; }
        public ShapeSheet.CellData<int> DisplayLevel { get; set; } // new in visio 2010
        public ShapeSheet.CellData<int> Relationships { get; set; } // new in visio 2010

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.ConFixedCode, this.ConFixedCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ConLineJumpCode, this.ConLineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ConLineJumpDirX, this.ConLineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ConLineJumpDirY, this.ConLineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ConLineJumpStyle, this.ConLineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ConLineRouteExt, this.ConLineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeFixedCode, this.ShapeFixedCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePermeablePlace, this.ShapePermeablePlace.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePermeableX, this.ShapePermeableX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePermeableY, this.ShapePermeableY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePlaceFlip, this.ShapePlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePlaceStyle, this.ShapePlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePlowCode, this.ShapePlowCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeRouteStyle, this.ShapeRouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeSplit, this.ShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeSplittable, this.ShapeSplittable.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DisplayLevel, this.DisplayLevel.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Relationships, this.Relationships.Formula);
            }
        }


        public static IList<ShapeLayoutCells> GetCells(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var query = ShapeLayoutCells.get_query();
            return CellGroup._GetCells<ShapeLayoutCells, double>(page, shapeids, query, query.GetCells);
        }

        public static ShapeLayoutCells GetCells(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = ShapeLayoutCells.get_query();
            return CellGroup._GetCells<ShapeLayoutCells, double>(shape, query, query.GetCells);
        }

        private static ShapeLayoutCellQuery _mCellQuery;
        private static ShapeLayoutCellQuery get_query()
        {
            ShapeLayoutCells._mCellQuery = ShapeLayoutCells._mCellQuery ?? new ShapeLayoutCellQuery();
            return ShapeLayoutCells._mCellQuery;
        }

        class ShapeLayoutCellQuery : CellQuery
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
                this.ConFixedCode = this.AddCell(ShapeSheet.SRCConstants.ConFixedCode, "ConFixedCode");
                this.ConLineJumpCode = this.AddCell(ShapeSheet.SRCConstants.ConLineJumpCode, "ConLineJumpCode");
                this.ConLineJumpDirX = this.AddCell(ShapeSheet.SRCConstants.ConLineJumpDirX, "ConLineJumpDirX");
                this.ConLineJumpDirY = this.AddCell(ShapeSheet.SRCConstants.ConLineJumpDirY, "ConLineJumpDirY");
                this.ConLineJumpStyle = this.AddCell(ShapeSheet.SRCConstants.ConLineJumpStyle, "ConLineJumpStyle");
                this.ConLineRouteExt = this.AddCell(ShapeSheet.SRCConstants.ConLineRouteExt, "ConLineRouteExt");
                this.ShapeFixedCode = this.AddCell(ShapeSheet.SRCConstants.ShapeFixedCode, "ShapeFixedCode");
                this.ShapePermeablePlace = this.AddCell(ShapeSheet.SRCConstants.ShapePermeablePlace, "ShapePermeablePlace");
                this.ShapePermeableX = this.AddCell(ShapeSheet.SRCConstants.ShapePermeableX, "ShapePermeableX");
                this.ShapePermeableY = this.AddCell(ShapeSheet.SRCConstants.ShapePermeableY, "ShapePermeableY");
                this.ShapePlaceFlip = this.AddCell(ShapeSheet.SRCConstants.ShapePlaceFlip, "ShapePlaceFlip");
                this.ShapePlaceStyle = this.AddCell(ShapeSheet.SRCConstants.ShapePlaceStyle, "ShapePlaceStyle");
                this.ShapePlowCode = this.AddCell(ShapeSheet.SRCConstants.ShapePlowCode, "ShapePlowCode");
                this.ShapeRouteStyle = this.AddCell(ShapeSheet.SRCConstants.ShapeRouteStyle, "ShapeRouteStyle");
                this.ShapeSplit = this.AddCell(ShapeSheet.SRCConstants.ShapeSplit, "ShapeSplit");
                this.ShapeSplittable = this.AddCell(ShapeSheet.SRCConstants.ShapeSplittable, "ShapeSplittable");
                this.DisplayLevel = this.AddCell(ShapeSheet.SRCConstants.DisplayLevel, "DisplayLevel");
                this.Relationships = this.AddCell(ShapeSheet.SRCConstants.Relationships, "Relationships");

            }

            public ShapeLayoutCells GetCells(IList<ShapeSheet.CellData<double>> row)
            {
                var cells = new ShapeLayoutCells();
                cells.ConFixedCode = row[this.ConFixedCode].ToInt();
                cells.ConLineJumpCode = row[this.ConLineJumpCode].ToInt();
                cells.ConLineJumpDirX = row[this.ConLineJumpDirX].ToInt();
                cells.ConLineJumpDirY = row[this.ConLineJumpDirY].ToInt();
                cells.ConLineJumpStyle = row[this.ConLineJumpStyle].ToInt();
                cells.ConLineRouteExt = row[this.ConLineRouteExt].ToInt();
                cells.ShapeFixedCode = row[this.ShapeFixedCode].ToInt();
                cells.ShapePermeablePlace = row[this.ShapePermeablePlace].ToInt();
                cells.ShapePermeableX = row[this.ShapePermeableX].ToInt();
                cells.ShapePermeableY = row[this.ShapePermeableY].ToInt();
                cells.ShapePlaceFlip = row[this.ShapePlaceFlip].ToInt();
                cells.ShapePlaceStyle = row[this.ShapePlaceStyle].ToInt();
                cells.ShapePlowCode = row[this.ShapePlowCode].ToInt();
                cells.ShapeRouteStyle = row[this.ShapeRouteStyle].ToInt();
                cells.ShapeSplit = row[this.ShapeSplit].ToInt();
                cells.ShapeSplittable = row[this.ShapeSplittable].ToInt();
                cells.DisplayLevel = row[this.DisplayLevel].ToInt();
                cells.Relationships = row[this.Relationships].ToInt();
                return cells;
            }
        }

    }
}