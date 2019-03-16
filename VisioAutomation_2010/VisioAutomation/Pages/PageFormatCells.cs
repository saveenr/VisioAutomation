using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageFormatCells : VASS.CellGroups.CellGroup
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

        public override IEnumerable<VASS.CellGroups.CellMetadataItem> CellMetadata
        {
            get
            {


                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.DrawingScale), VASS.SrcConstants.PageDrawingScale, this.DrawingScale);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.DrawingScaleType), VASS.SrcConstants.PageDrawingScaleType, this.DrawingScaleType);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.DrawingSizeType), VASS.SrcConstants.PageDrawingSizeType, this.DrawingSizeType);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.InhibitSnap), VASS.SrcConstants.PageInhibitSnap, this.InhibitSnap);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Height), VASS.SrcConstants.PageHeight, this.Height);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Scale), VASS.SrcConstants.PageScale, this.Scale);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Width), VASS.SrcConstants.PageWidth, this.Width);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.ShadowObliqueAngle), VASS.SrcConstants.PageShadowObliqueAngle, this.ShadowObliqueAngle);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.ShadowOffsetX), VASS.SrcConstants.PageShadowOffsetX, this.ShadowOffsetX);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.ShadowOffsetY), VASS.SrcConstants.PageShadowOffsetY, this.ShadowOffsetY);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.ShadowScaleFactor), VASS.SrcConstants.PageShadowScaleFactor, this.ShadowScaleFactor);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.ShadowType), VASS.SrcConstants.PageShadowType, this.ShadowType);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.UIVisibility), VASS.SrcConstants.PageUIVisibility, this.UIVisibility);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.DrawingResizeType), VASS.SrcConstants.PageDrawingResizeType, this.DrawingResizeType);
            }
        }



        public static PageFormatCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = PageFormatCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageFormatCellsBuilder> PageFormatCells_lazy_builder = new System.Lazy<PageFormatCellsBuilder>();

        class PageFormatCellsBuilder : VASS.CellGroups.CellGroupBuilder<PageFormatCells>
        {
            public PageFormatCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override PageFormatCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new PageFormatCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.DrawingScale = getcellvalue(nameof(PageFormatCells.DrawingScale));
                cells.DrawingScaleType = getcellvalue(nameof(PageFormatCells.DrawingScaleType));
                cells.DrawingSizeType = getcellvalue(nameof(PageFormatCells.DrawingSizeType));
                cells.InhibitSnap = getcellvalue(nameof(PageFormatCells.InhibitSnap));
                cells.Height = getcellvalue(nameof(PageFormatCells.Height));
                cells.Scale = getcellvalue(nameof(PageFormatCells.Scale));
                cells.Width = getcellvalue(nameof(PageFormatCells.Width));
                cells.ShadowObliqueAngle = getcellvalue(nameof(PageFormatCells.ShadowObliqueAngle));
                cells.ShadowOffsetX = getcellvalue(nameof(PageFormatCells.ShadowOffsetX));
                cells.ShadowOffsetY = getcellvalue(nameof(PageFormatCells.ShadowOffsetY));
                cells.ShadowScaleFactor = getcellvalue(nameof(PageFormatCells.ShadowScaleFactor));
                cells.ShadowType = getcellvalue(nameof(PageFormatCells.ShadowType));
                cells.UIVisibility = getcellvalue(nameof(PageFormatCells.UIVisibility));
                cells.DrawingResizeType = getcellvalue(nameof(PageFormatCells.DrawingResizeType));

                return cells;
            }
        }

    }
}