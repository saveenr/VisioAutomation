using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageFormatCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral DrawingScale { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DrawingScaleType { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DrawingSizeType { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral InhibitSnap { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Height { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Scale { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Width { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShadowObliqueAngle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShadowOffsetX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShadowOffsetY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShadowScaleFactor { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShadowType { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral UIVisibility { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            { 
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageDrawingScale, this.DrawingScale.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageDrawingScaleType, this.DrawingScaleType.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageDrawingSizeType, this.DrawingSizeType.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageInhibitSnap, this.InhibitSnap.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageHeight, this.Height.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageScale, this.Scale.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageWidth, this.Width.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageShadowObliqueAngle, this.ShadowObliqueAngle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageShadowOffsetX, this.ShadowOffsetX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageShadowOffsetY, this.ShadowOffsetY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageShadowScaleFactor, this.ShadowScaleFactor.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageShadowType, this.ShadowType.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageUIVisibility, this.UIVisibility.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageDrawingResizeType, this.DrawingResizeType.Value);
            }
        }

        public static PageFormatCells GetFormulas(IVisio.Shape shape)
        {
            var query = PageFormatCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Formula);
        }

        public static PageFormatCells GetResults(IVisio.Shape shape)
        {
            var query = PageFormatCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Result);
        }

        private static readonly System.Lazy<PageFormatCellsReader> lazy_query = new System.Lazy<PageFormatCellsReader>();

        class PageFormatCellsReader : ReaderSingleRow<VisioAutomation.Pages.PageFormatCells>
        {
            public CellColumn DrawingScale { get; set; }
            public CellColumn DrawingScaleType { get; set; }
            public CellColumn DrawingSizeType { get; set; }
            public CellColumn InhibitSnap { get; set; }
            public CellColumn Height { get; set; }
            public CellColumn Scale { get; set; }
            public CellColumn Width { get; set; }
            public CellColumn ShadowObliqueAngle { get; set; }
            public CellColumn ShadowOffsetX { get; set; }
            public CellColumn ShadowOffsetY { get; set; }
            public CellColumn ShadowScaleFactor { get; set; }
            public CellColumn ShadowType { get; set; }
            public CellColumn UIVisibility { get; set; }
            public CellColumn DrawingResizeType { get; set; }

            public PageFormatCellsReader()
            {
                this.DrawingScale = this.query.Columns.Add(SrcConstants.PageDrawingScale, nameof(SrcConstants.PageDrawingScale));
                this.DrawingScaleType = this.query.Columns.Add(SrcConstants.PageDrawingScaleType, nameof(SrcConstants.PageDrawingScaleType));
                this.DrawingSizeType = this.query.Columns.Add(SrcConstants.PageDrawingSizeType, nameof(SrcConstants.PageDrawingSizeType));
                this.InhibitSnap = this.query.Columns.Add(SrcConstants.PageInhibitSnap, nameof(SrcConstants.PageInhibitSnap));
                this.Height = this.query.Columns.Add(SrcConstants.PageHeight, nameof(SrcConstants.PageHeight));
                this.Scale = this.query.Columns.Add(SrcConstants.PageScale, nameof(SrcConstants.PageScale));
                this.Width = this.query.Columns.Add(SrcConstants.PageWidth, nameof(SrcConstants.PageWidth));
                this.ShadowObliqueAngle = this.query.Columns.Add(SrcConstants.PageShadowObliqueAngle, nameof(SrcConstants.PageShadowObliqueAngle));
                this.ShadowOffsetX = this.query.Columns.Add(SrcConstants.PageShadowOffsetX, nameof(SrcConstants.PageShadowOffsetX));
                this.ShadowOffsetY = this.query.Columns.Add(SrcConstants.PageShadowOffsetY, nameof(SrcConstants.PageShadowOffsetY));
                this.ShadowScaleFactor = this.query.Columns.Add(SrcConstants.PageShadowScaleFactor, nameof(SrcConstants.PageShadowScaleFactor));
                this.ShadowType = this.query.Columns.Add(SrcConstants.PageShadowType, nameof(SrcConstants.PageShadowType));
                this.UIVisibility = this.query.Columns.Add(SrcConstants.PageUIVisibility, nameof(SrcConstants.PageUIVisibility));
                this.DrawingResizeType = this.query.Columns.Add(SrcConstants.PageDrawingResizeType, nameof(SrcConstants.PageDrawingResizeType));
            }

            public override Pages.PageFormatCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new Pages.PageFormatCells();
                cells.DrawingScale = row[this.DrawingScale];
                cells.DrawingScaleType = row[this.DrawingScaleType];
                cells.DrawingSizeType = row[this.DrawingSizeType];
                cells.InhibitSnap = row[this.InhibitSnap];
                cells.Height = row[this.Height];
                cells.Scale = row[this.Scale];
                cells.Width = row[this.Width];
                cells.ShadowObliqueAngle = row[this.ShadowObliqueAngle];
                cells.ShadowOffsetX = row[this.ShadowOffsetX];
                cells.ShadowOffsetY = row[this.ShadowOffsetY];
                cells.ShadowScaleFactor = row[this.ShadowScaleFactor];
                cells.ShadowType = row[this.ShadowType];
                cells.UIVisibility = row[this.UIVisibility];
                cells.DrawingResizeType = row[this.DrawingResizeType];
                return cells;
            }
        }

    }
}