using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageFormatCells : VASS.CellGroups.CellGroupBase
    {
        public VASS.CellValueLiteral DrawingScale { get; set; }
        public VASS.CellValueLiteral DrawingScaleType { get; set; }
        public VASS.CellValueLiteral DrawingSizeType { get; set; }
        public VASS.CellValueLiteral InhibitSnap { get; set; }
        public VASS.CellValueLiteral Height { get; set; }
        public VASS.CellValueLiteral Scale { get; set; }
        public VASS.CellValueLiteral Width { get; set; }
        public VASS.CellValueLiteral ShadowObliqueAngle { get; set; }
        public VASS.CellValueLiteral ShadowOffsetX { get; set; }
        public VASS.CellValueLiteral ShadowOffsetY { get; set; }
        public VASS.CellValueLiteral ShadowScaleFactor { get; set; }
        public VASS.CellValueLiteral ShadowType { get; set; }
        public VASS.CellValueLiteral UIVisibility { get; set; }
        public VASS.CellValueLiteral DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<VASS.CellGroups.SrcValuePair> SrcValuePairs
        {
            get
            { 
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageDrawingScale, this.DrawingScale);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageDrawingScaleType, this.DrawingScaleType);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageDrawingSizeType, this.DrawingSizeType);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageInhibitSnap, this.InhibitSnap);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageHeight, this.Height);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageScale, this.Scale);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageWidth, this.Width);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageShadowObliqueAngle, this.ShadowObliqueAngle);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageShadowOffsetX, this.ShadowOffsetX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageShadowOffsetY, this.ShadowOffsetY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageShadowScaleFactor, this.ShadowScaleFactor);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageShadowType, this.ShadowType);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageUIVisibility, this.UIVisibility);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageDrawingResizeType, this.DrawingResizeType);
            }
        }

        public static PageFormatCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(shape, type);
        }

        private static readonly System.Lazy<PageFormatCellsReader> lazy_query = new System.Lazy<PageFormatCellsReader>();

        class PageFormatCellsReader : VASS.CellGroups.ReaderSingleRow<PageFormatCells>
        {
            public VASS.Query.CellColumn DrawingScale { get; set; }
            public VASS.Query.CellColumn DrawingScaleType { get; set; }
            public VASS.Query.CellColumn DrawingSizeType { get; set; }
            public VASS.Query.CellColumn InhibitSnap { get; set; }
            public VASS.Query.CellColumn Height { get; set; }
            public VASS.Query.CellColumn Scale { get; set; }
            public VASS.Query.CellColumn Width { get; set; }
            public VASS.Query.CellColumn ShadowObliqueAngle { get; set; }
            public VASS.Query.CellColumn ShadowOffsetX { get; set; }
            public VASS.Query.CellColumn ShadowOffsetY { get; set; }
            public VASS.Query.CellColumn ShadowScaleFactor { get; set; }
            public VASS.Query.CellColumn ShadowType { get; set; }
            public VASS.Query.CellColumn UIVisibility { get; set; }
            public VASS.Query.CellColumn DrawingResizeType { get; set; }

            public PageFormatCellsReader()
            {
                this.DrawingScale = this.query.Columns.Add(VASS.SrcConstants.PageDrawingScale, nameof(this.DrawingScale));
                this.DrawingScaleType = this.query.Columns.Add(VASS.SrcConstants.PageDrawingScaleType, nameof(this.DrawingScaleType));
                this.DrawingSizeType = this.query.Columns.Add(VASS.SrcConstants.PageDrawingSizeType, nameof(this.DrawingSizeType));
                this.InhibitSnap = this.query.Columns.Add(VASS.SrcConstants.PageInhibitSnap, nameof(this.InhibitSnap));
                this.Height = this.query.Columns.Add(VASS.SrcConstants.PageHeight, nameof(this.Height));
                this.Scale = this.query.Columns.Add(VASS.SrcConstants.PageScale, nameof(this.Scale));
                this.Width = this.query.Columns.Add(VASS.SrcConstants.PageWidth, nameof(this.Width));
                this.ShadowObliqueAngle = this.query.Columns.Add(VASS.SrcConstants.PageShadowObliqueAngle, nameof(this.ShadowObliqueAngle));
                this.ShadowOffsetX = this.query.Columns.Add(VASS.SrcConstants.PageShadowOffsetX, nameof(this.ShadowOffsetX));
                this.ShadowOffsetY = this.query.Columns.Add(VASS.SrcConstants.PageShadowOffsetY, nameof(this.ShadowOffsetY));
                this.ShadowScaleFactor = this.query.Columns.Add(VASS.SrcConstants.PageShadowScaleFactor, nameof(this.ShadowScaleFactor));
                this.ShadowType = this.query.Columns.Add(VASS.SrcConstants.PageShadowType, nameof(this.ShadowType));
                this.UIVisibility = this.query.Columns.Add(VASS.SrcConstants.PageUIVisibility, nameof(this.UIVisibility));
                this.DrawingResizeType = this.query.Columns.Add(VASS.SrcConstants.PageDrawingResizeType, nameof(this.DrawingResizeType));
            }

            public override PageFormatCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
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