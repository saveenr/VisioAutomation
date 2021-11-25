using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;

using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageFormatCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue DrawingScale { get; set; }
        public VisioAutomation.Core.CellValue DrawingScaleType { get; set; }
        public VisioAutomation.Core.CellValue DrawingSizeType { get; set; }
        public VisioAutomation.Core.CellValue InhibitSnap { get; set; }
        public VisioAutomation.Core.CellValue Height { get; set; }
        public VisioAutomation.Core.CellValue Scale { get; set; }
        public VisioAutomation.Core.CellValue Width { get; set; }
        public VisioAutomation.Core.CellValue ShadowObliqueAngle { get; set; }
        public VisioAutomation.Core.CellValue ShadowOffsetX { get; set; }
        public VisioAutomation.Core.CellValue ShadowOffsetY { get; set; }
        public VisioAutomation.Core.CellValue ShadowScaleFactor { get; set; }
        public VisioAutomation.Core.CellValue ShadowType { get; set; }
        public VisioAutomation.Core.CellValue UIVisibility { get; set; }
        public VisioAutomation.Core.CellValue DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.DrawingScale), VisioAutomation.Core.SrcConstants.PageDrawingScale, this.DrawingScale);
            yield return this.Create(nameof(this.DrawingScaleType), VisioAutomation.Core.SrcConstants.PageDrawingScaleType,
                this.DrawingScaleType);
            yield return this.Create(nameof(this.DrawingSizeType), VisioAutomation.Core.SrcConstants.PageDrawingSizeType, this.DrawingSizeType);
            yield return this.Create(nameof(this.InhibitSnap), VisioAutomation.Core.SrcConstants.PageInhibitSnap, this.InhibitSnap);
            yield return this.Create(nameof(this.Height), VisioAutomation.Core.SrcConstants.PageHeight, this.Height);
            yield return this.Create(nameof(this.Scale), VisioAutomation.Core.SrcConstants.PageScale, this.Scale);
            yield return this.Create(nameof(this.Width), VisioAutomation.Core.SrcConstants.PageWidth, this.Width);
            yield return this.Create(nameof(this.ShadowObliqueAngle), VisioAutomation.Core.SrcConstants.PageShadowObliqueAngle,
                this.ShadowObliqueAngle);
            yield return this.Create(nameof(this.ShadowOffsetX), VisioAutomation.Core.SrcConstants.PageShadowOffsetX, this.ShadowOffsetX);
            yield return this.Create(nameof(this.ShadowOffsetY), VisioAutomation.Core.SrcConstants.PageShadowOffsetY, this.ShadowOffsetY);
            yield return this.Create(nameof(this.ShadowScaleFactor), VisioAutomation.Core.SrcConstants.PageShadowScaleFactor,
                this.ShadowScaleFactor);
            yield return this.Create(nameof(this.ShadowType), VisioAutomation.Core.SrcConstants.PageShadowType, this.ShadowType);
            yield return this.Create(nameof(this.UIVisibility), VisioAutomation.Core.SrcConstants.PageUIVisibility, this.UIVisibility);
            yield return this.Create(nameof(this.DrawingResizeType), VisioAutomation.Core.SrcConstants.PageDrawingResizeType,
                this.DrawingResizeType);
        }


        public static PageFormatCells GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
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

            public override PageFormatCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new PageFormatCells();
                var getcellvalue = VASS.CellGroups.CellGroup.row_to_cellgroup(row, cols);

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