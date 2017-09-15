using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageFormatCells : CellGroupSingleRow
    {
        public CellValueLiteral DrawingScale { get; set; }
        public CellValueLiteral DrawingScaleType { get; set; }
        public CellValueLiteral DrawingSizeType { get; set; }
        public CellValueLiteral InhibitSnap { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral Scale { get; set; }
        public CellValueLiteral Width { get; set; }
        public CellValueLiteral ShadowObliqueAngle { get; set; }
        public CellValueLiteral ShadowOffsetX { get; set; }
        public CellValueLiteral ShadowOffsetY { get; set; }
        public CellValueLiteral ShadowScaleFactor { get; set; }
        public CellValueLiteral ShadowType { get; set; }
        public CellValueLiteral UIVisibility { get; set; }
        public CellValueLiteral DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            { 
                yield return SrcValuePair.Create(SrcConstants.PageDrawingScale, this.DrawingScale);
                yield return SrcValuePair.Create(SrcConstants.PageDrawingScaleType, this.DrawingScaleType);
                yield return SrcValuePair.Create(SrcConstants.PageDrawingSizeType, this.DrawingSizeType);
                yield return SrcValuePair.Create(SrcConstants.PageInhibitSnap, this.InhibitSnap);
                yield return SrcValuePair.Create(SrcConstants.PageHeight, this.Height);
                yield return SrcValuePair.Create(SrcConstants.PageScale, this.Scale);
                yield return SrcValuePair.Create(SrcConstants.PageWidth, this.Width);
                yield return SrcValuePair.Create(SrcConstants.PageShadowObliqueAngle, this.ShadowObliqueAngle);
                yield return SrcValuePair.Create(SrcConstants.PageShadowOffsetX, this.ShadowOffsetX);
                yield return SrcValuePair.Create(SrcConstants.PageShadowOffsetY, this.ShadowOffsetY);
                yield return SrcValuePair.Create(SrcConstants.PageShadowScaleFactor, this.ShadowScaleFactor);
                yield return SrcValuePair.Create(SrcConstants.PageShadowType, this.ShadowType);
                yield return SrcValuePair.Create(SrcConstants.PageUIVisibility, this.UIVisibility);
                yield return SrcValuePair.Create(SrcConstants.PageDrawingResizeType, this.DrawingResizeType);
            }
        }

        public static PageFormatCells GetCells(IVisio.Shape shape, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetCells(shape, cvt);
        }

        private static readonly System.Lazy<PageFormatCellsReader> lazy_query = new System.Lazy<PageFormatCellsReader>();

        class PageFormatCellsReader : ReaderSingleRow<PageFormatCells>
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
                this.DrawingScale = this.query.Columns.Add(SrcConstants.PageDrawingScale, nameof(this.DrawingScale));
                this.DrawingScaleType = this.query.Columns.Add(SrcConstants.PageDrawingScaleType, nameof(this.DrawingScaleType));
                this.DrawingSizeType = this.query.Columns.Add(SrcConstants.PageDrawingSizeType, nameof(this.DrawingSizeType));
                this.InhibitSnap = this.query.Columns.Add(SrcConstants.PageInhibitSnap, nameof(this.InhibitSnap));
                this.Height = this.query.Columns.Add(SrcConstants.PageHeight, nameof(this.Height));
                this.Scale = this.query.Columns.Add(SrcConstants.PageScale, nameof(this.Scale));
                this.Width = this.query.Columns.Add(SrcConstants.PageWidth, nameof(this.Width));
                this.ShadowObliqueAngle = this.query.Columns.Add(SrcConstants.PageShadowObliqueAngle, nameof(this.ShadowObliqueAngle));
                this.ShadowOffsetX = this.query.Columns.Add(SrcConstants.PageShadowOffsetX, nameof(this.ShadowOffsetX));
                this.ShadowOffsetY = this.query.Columns.Add(SrcConstants.PageShadowOffsetY, nameof(this.ShadowOffsetY));
                this.ShadowScaleFactor = this.query.Columns.Add(SrcConstants.PageShadowScaleFactor, nameof(this.ShadowScaleFactor));
                this.ShadowType = this.query.Columns.Add(SrcConstants.PageShadowType, nameof(this.ShadowType));
                this.UIVisibility = this.query.Columns.Add(SrcConstants.PageUIVisibility, nameof(this.UIVisibility));
                this.DrawingResizeType = this.query.Columns.Add(SrcConstants.PageDrawingResizeType, nameof(this.DrawingResizeType));
            }

            public override PageFormatCells CellDataToCellGroup(Utilities.ArraySegment<string> row)
            {
                var cells = new PageFormatCells();
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