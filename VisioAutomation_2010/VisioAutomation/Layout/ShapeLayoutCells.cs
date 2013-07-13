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

        private static ShapeLayoutQuery m_query;
        private static ShapeLayoutQuery get_query()
        {
            m_query = m_query ?? new ShapeLayoutQuery();
            return m_query;
        }

        class ShapeLayoutQuery : VA.ShapeSheet.Query.QueryEx
        {
            public int ConFixedCode { get; set; }
            public int ConLineJumpCode { get; set; }
            public int ConLineJumpDirX { get; set; }
            public int ConLineJumpDirY { get; set; }
            public int ConLineJumpStyle { get; set; }
            public int ConLineRouteExt { get; set; }
            public int ShapeFixedCode { get; set; }
            public int ShapePermeablePlace { get; set; }
            public int ShapePermeableX { get; set; }
            public int ShapePermeableY { get; set; }
            public int ShapePlaceFlip { get; set; }
            public int ShapePlaceStyle { get; set; }
            public int ShapePlowCode { get; set; }
            public int ShapeRouteStyle { get; set; }
            public int ShapeSplit { get; set; }
            public int ShapeSplittable { get; set; }
            public int DisplayLevel { get; set; }
            public int Relationships { get; set; }

            public ShapeLayoutQuery() :
                base()
            {
                this.ConFixedCode = this.AddColumn2(VA.ShapeSheet.SRCConstants.ConFixedCode, "ConFixedCode");
                this.ConLineJumpCode = this.AddColumn2(VA.ShapeSheet.SRCConstants.ConLineJumpCode, "ConLineJumpCode");
                this.ConLineJumpDirX = this.AddColumn2(VA.ShapeSheet.SRCConstants.ConLineJumpDirX, "ConLineJumpDirX");
                this.ConLineJumpDirY = this.AddColumn2(VA.ShapeSheet.SRCConstants.ConLineJumpDirY, "ConLineJumpDirY");
                this.ConLineJumpStyle = this.AddColumn2(VA.ShapeSheet.SRCConstants.ConLineJumpStyle, "ConLineJumpStyle");
                this.ConLineRouteExt = this.AddColumn2(VA.ShapeSheet.SRCConstants.ConLineRouteExt, "ConLineRouteExt");
                this.ShapeFixedCode = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShapeFixedCode, "ShapeFixedCode");
                this.ShapePermeablePlace = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShapePermeablePlace, "ShapePermeablePlace");
                this.ShapePermeableX = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShapePermeableX, "ShapePermeableX");
                this.ShapePermeableY = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShapePermeableY, "ShapePermeableY");
                this.ShapePlaceFlip = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShapePlaceFlip, "ShapePlaceFlip");
                this.ShapePlaceStyle = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShapePlaceStyle, "ShapePlaceStyle");
                this.ShapePlowCode = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShapePlowCode, "ShapePlowCode");
                this.ShapeRouteStyle = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShapeRouteStyle, "ShapeRouteStyle");
                this.ShapeSplit = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShapeSplit, "ShapeSplit");
                this.ShapeSplittable = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShapeSplittable, "ShapeSplittable");
                this.DisplayLevel= this.AddColumn2(VA.ShapeSheet.SRCConstants.DisplayLevel, "DisplayLevel");
                this.Relationships = this.AddColumn2(VA.ShapeSheet.SRCConstants.Relationships, "Relationships");
            }

            public ShapeLayoutCells GetCells(ExQueryResult<CellData<double>> data_for_shape)
            {
                var row = data_for_shape.Cells;
                var cells = new ShapeLayoutCells();
                cells.ConFixedCode = row[ConFixedCode].ToInt();
                cells.ConLineJumpCode = row[ConLineJumpCode].ToInt();
                cells.ConLineJumpDirX = row[ConLineJumpDirX].ToInt();
                cells.ConLineJumpDirY = row[ConLineJumpDirY].ToInt();
                cells.ConLineJumpStyle = row[ConLineJumpStyle].ToInt();
                cells.ConLineRouteExt = row[ConLineRouteExt].ToInt();
                cells.ShapeFixedCode = row[ShapeFixedCode].ToInt();
                cells.ShapePermeablePlace = row[ShapePermeablePlace].ToInt();
                cells.ShapePermeableX = row[ShapePermeableX].ToInt();
                cells.ShapePermeableY = row[ShapePermeableY].ToInt();
                cells.ShapePlaceFlip = row[ShapePlaceFlip].ToInt();
                cells.ShapePlaceStyle = row[ShapePlaceStyle].ToInt();
                cells.ShapePlowCode = row[ShapePlowCode].ToInt();
                cells.ShapeRouteStyle = row[ShapeRouteStyle].ToInt();
                cells.ShapeSplit = row[ShapeSplit].ToInt();
                cells.ShapeSplittable = row[ShapeSplittable].ToInt();
                cells.DisplayLevel = row[DisplayLevel].ToInt();
                cells.Relationships = row[Relationships].ToInt();
                return cells;
            }
        }

    }
}