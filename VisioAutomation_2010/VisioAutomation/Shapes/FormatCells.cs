using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using VAQUERY=VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class FormatCells : ShapeSheet.CellGroups.CellGroup
    {
        public ShapeSheet.CellData<int> FillBkgnd { get; set; }
        public ShapeSheet.CellData<double> FillBkgndTrans { get; set; }
        public ShapeSheet.CellData<int> FillForegnd { get; set; }
        public ShapeSheet.CellData<double> FillForegndTrans { get; set; }
        public ShapeSheet.CellData<int> FillPattern { get; set; }
        public ShapeSheet.CellData<double> ShapeShdwObliqueAngle { get; set; }
        public ShapeSheet.CellData<double> ShapeShdwOffsetX { get; set; }
        public ShapeSheet.CellData<double> ShapeShdwOffsetY { get; set; }
        public ShapeSheet.CellData<double> ShapeShdwScaleFactor { get; set; }
        public ShapeSheet.CellData<int> ShapeShdwType { get; set; }
        public ShapeSheet.CellData<int> ShdwBkgnd { get; set; }
        public ShapeSheet.CellData<double> ShdwBkgndTrans { get; set; }
        public ShapeSheet.CellData<int> ShdwForegnd { get; set; }
        public ShapeSheet.CellData<double> ShdwForegndTrans { get; set; }
        public ShapeSheet.CellData<int> ShdwPattern { get; set; }
        public ShapeSheet.CellData<int> BeginArrow { get; set; }
        public ShapeSheet.CellData<double> BeginArrowSize { get; set; }
        public ShapeSheet.CellData<int> EndArrow { get; set; }
        public ShapeSheet.CellData<double> EndArrowSize { get; set; }
        public ShapeSheet.CellData<int> LineCap { get; set; }
        public ShapeSheet.CellData<int> LineColor { get; set; }
        public ShapeSheet.CellData<double> LineColorTrans { get; set; }
        public ShapeSheet.CellData<int> LinePattern { get; set; }
        public ShapeSheet.CellData<double> LineWeight { get; set; }
        public ShapeSheet.CellData<double> Rounding { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.FillBkgnd, this.FillBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.FillBkgndTrans, this.FillBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.FillForegnd, this.FillForegnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.FillForegndTrans, this.FillForegndTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.FillPattern, this.FillPattern.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, this.ShapeShdwObliqueAngle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeShdwOffsetX, this.ShapeShdwOffsetX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeShdwOffsetY, this.ShapeShdwOffsetY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeShdwScaleFactor, this.ShapeShdwScaleFactor.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeShdwType, this.ShapeShdwType.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwBkgnd, this.ShdwBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwBkgndTrans, this.ShdwBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwForegnd, this.ShdwForegnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwForegndTrans, this.ShdwForegndTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwPattern, this.ShdwPattern.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.BeginArrow, this.BeginArrow.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.BeginArrowSize, this.BeginArrowSize.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.EndArrow, this.EndArrow.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.EndArrowSize, this.EndArrowSize.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineCap, this.LineCap.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineColor, this.LineColor.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineColorTrans, this.LineColorTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LinePattern, this.LinePattern.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineWeight, this.LineWeight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Rounding, this.Rounding.Formula);
            }
        }


        public static IList<FormatCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = FormatCells.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroup._GetCells<FormatCells, double>(page, shapeids, query, query.GetCells);
        }

        public static FormatCells GetCells(IVisio.Shape shape)
        {
            var query = FormatCells.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroup._GetCells<FormatCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeFormatCellQuery> lazy_query = new System.Lazy<ShapeFormatCellQuery>();

        class ShapeFormatCellQuery : VAQUERY.CellQuery
        {
            public VAQUERY.CellColumn FillBkgnd { get; set; }
            public VAQUERY.CellColumn FillBkgndTrans { get; set; }
            public VAQUERY.CellColumn FillForegnd { get; set; }
            public VAQUERY.CellColumn FillForegndTrans { get; set; }
            public VAQUERY.CellColumn FillPattern { get; set; }
            public VAQUERY.CellColumn ShapeShdwObliqueAngle { get; set; }
            public VAQUERY.CellColumn ShapeShdwOffsetX { get; set; }
            public VAQUERY.CellColumn ShapeShdwOffsetY { get; set; }
            public VAQUERY.CellColumn ShapeShdwScaleFactor { get; set; }
            public VAQUERY.CellColumn ShapeShdwType { get; set; }
            public VAQUERY.CellColumn ShdwBkgnd { get; set; }
            public VAQUERY.CellColumn ShdwBkgndTrans { get; set; }
            public VAQUERY.CellColumn ShdwForegnd { get; set; }
            public VAQUERY.CellColumn ShdwForegndTrans { get; set; }
            public VAQUERY.CellColumn ShdwPattern { get; set; }
            public VAQUERY.CellColumn BeginArrow { get; set; }
            public VAQUERY.CellColumn BeginArrowSize { get; set; }
            public VAQUERY.CellColumn EndArrow { get; set; }
            public VAQUERY.CellColumn EndArrowSize { get; set; }
            public VAQUERY.CellColumn LineColor { get; set; }
            public VAQUERY.CellColumn LineCap { get; set; }
            public VAQUERY.CellColumn LineColorTrans { get; set; }
            public VAQUERY.CellColumn LinePattern { get; set; }
            public VAQUERY.CellColumn LineWeight { get; set; }
            public VAQUERY.CellColumn Rounding { get; set; }

            public ShapeFormatCellQuery()
            {
                this.FillBkgnd = this.AddCell(ShapeSheet.SRCConstants.FillBkgnd, "FillBkgnd");
                this.FillBkgndTrans = this.AddCell(ShapeSheet.SRCConstants.FillBkgndTrans, "FillBkgndTrans");
                this.FillForegnd = this.AddCell(ShapeSheet.SRCConstants.FillForegnd, "FillForegnd");
                this.FillForegndTrans = this.AddCell(ShapeSheet.SRCConstants.FillForegndTrans, "FillForegndTrans");
                this.FillPattern = this.AddCell(ShapeSheet.SRCConstants.FillPattern, "FillPattern");
                this.ShapeShdwObliqueAngle = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, "ShapeShdwObliqueAngle");
                this.ShapeShdwOffsetX = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwOffsetX, "ShapeShdwOffsetX");
                this.ShapeShdwOffsetY = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwOffsetY, "ShapeShdwOffsetY");
                this.ShapeShdwScaleFactor = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwScaleFactor, "ShapeShdwScaleFactor");
                this.ShapeShdwType = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwType, "ShapeShdwType");
                this.ShdwBkgnd = this.AddCell(ShapeSheet.SRCConstants.ShdwBkgnd, "ShdwBkgnd");
                this.ShdwBkgndTrans = this.AddCell(ShapeSheet.SRCConstants.ShdwBkgndTrans, "ShdwBkgndTrans");
                this.ShdwForegnd = this.AddCell(ShapeSheet.SRCConstants.ShdwForegnd, "ShdwForegnd");
                this.ShdwForegndTrans = this.AddCell(ShapeSheet.SRCConstants.ShdwForegndTrans, "ShdwForegndTrans");
                this.ShdwPattern = this.AddCell(ShapeSheet.SRCConstants.ShdwPattern, "ShdwPattern");

                this.BeginArrow = this.AddCell(ShapeSheet.SRCConstants.BeginArrow, "BeginArrow");
                this.BeginArrowSize = this.AddCell(ShapeSheet.SRCConstants.BeginArrowSize, "BeginArrowSize");
                this.EndArrow = this.AddCell(ShapeSheet.SRCConstants.EndArrow, "EndArrow");
                this.EndArrowSize = this.AddCell(ShapeSheet.SRCConstants.EndArrowSize, "EndArrowSize");
                this.LineColor = this.AddCell(ShapeSheet.SRCConstants.LineColor, "LineColor");
                this.LineCap = this.AddCell(ShapeSheet.SRCConstants.LineCap, "LineCap");
                this.LineColorTrans = this.AddCell(ShapeSheet.SRCConstants.LineColorTrans, "LineColorTrans");
                this.LinePattern = this.AddCell(ShapeSheet.SRCConstants.LinePattern, "LinePattern");
                this.LineWeight = this.AddCell(ShapeSheet.SRCConstants.LineWeight, "LineWeight");
                this.Rounding = this.AddCell(ShapeSheet.SRCConstants.Rounding, "Rounding");

            }

            public FormatCells GetCells(IList<ShapeSheet.CellData<double>> row)
            {
                var cells = new FormatCells();
                cells.FillBkgnd = row[ this.FillBkgnd].ToInt();
                cells.FillBkgndTrans = row[ this.FillBkgndTrans];
                cells.FillForegnd = row[ this.FillForegnd].ToInt();
                cells.FillForegndTrans = row[ this.FillForegndTrans];
                cells.FillPattern = row[ this.FillPattern].ToInt();
                cells.ShapeShdwObliqueAngle = row[ this.ShapeShdwObliqueAngle];
                cells.ShapeShdwOffsetX = row[ this.ShapeShdwOffsetX];
                cells.ShapeShdwOffsetY = row[ this.ShapeShdwOffsetY];
                cells.ShapeShdwScaleFactor = row[ this.ShapeShdwScaleFactor];
                cells.ShapeShdwType = row[ this.ShapeShdwType].ToInt();
                cells.ShdwBkgnd = row[ this.ShdwBkgnd].ToInt();
                cells.ShdwBkgndTrans = row[ this.ShdwBkgndTrans];
                cells.ShdwForegnd = row[ this.ShdwForegnd].ToInt();
                cells.ShdwForegndTrans = row[ this.ShdwForegndTrans];
                cells.ShdwPattern = row[ this.ShdwPattern].ToInt();
                cells.BeginArrow = row[ this.BeginArrow].ToInt();
                cells.BeginArrowSize = row[ this.BeginArrowSize];
                cells.EndArrow = row[ this.EndArrow].ToInt();
                cells.EndArrowSize = row[ this.EndArrowSize];
                cells.LineCap = row[ this.LineCap].ToInt();
                cells.LineColor = row[ this.LineColor].ToInt();
                cells.LineColorTrans = row[ this.LineColorTrans];
                cells.LinePattern = row[ this.LinePattern].ToInt();
                cells.LineWeight = row[ this.LineWeight];
                cells.Rounding = row[ this.Rounding];
                return cells;
            }

        }
    }
}

