using VisioAutomation.ShapeSheet.CellGroups;



namespace VisioAutomation.Pages
{
    public class PageFormatCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValue DrawingScale { get; set; }
        public VASS.CellValue DrawingScaleType { get; set; }
        public VASS.CellValue DrawingSizeType { get; set; }
        public VASS.CellValue InhibitSnap { get; set; }
        public VASS.CellValue Height { get; set; }
        public VASS.CellValue Scale { get; set; }
        public VASS.CellValue Width { get; set; }
        public VASS.CellValue ShadowObliqueAngle { get; set; }
        public VASS.CellValue ShadowOffsetX { get; set; }
        public VASS.CellValue ShadowOffsetY { get; set; }
        public VASS.CellValue ShadowScaleFactor { get; set; }
        public VASS.CellValue ShadowType { get; set; }
        public VASS.CellValue UIVisibility { get; set; }
        public VASS.CellValue DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.DrawingScale), VASS.SrcConstants.PageDrawingScale, this.DrawingScale);
            yield return this.Create(nameof(this.DrawingScaleType), VASS.SrcConstants.PageDrawingScaleType,
                this.DrawingScaleType);
            yield return this.Create(nameof(this.DrawingSizeType), VASS.SrcConstants.PageDrawingSizeType, this.DrawingSizeType);
            yield return this.Create(nameof(this.InhibitSnap), VASS.SrcConstants.PageInhibitSnap, this.InhibitSnap);
            yield return this.Create(nameof(this.Height), VASS.SrcConstants.PageHeight, this.Height);
            yield return this.Create(nameof(this.Scale), VASS.SrcConstants.PageScale, this.Scale);
            yield return this.Create(nameof(this.Width), VASS.SrcConstants.PageWidth, this.Width);
            yield return this.Create(nameof(this.ShadowObliqueAngle), VASS.SrcConstants.PageShadowObliqueAngle,
                this.ShadowObliqueAngle);
            yield return this.Create(nameof(this.ShadowOffsetX), VASS.SrcConstants.PageShadowOffsetX, this.ShadowOffsetX);
            yield return this.Create(nameof(this.ShadowOffsetY), VASS.SrcConstants.PageShadowOffsetY, this.ShadowOffsetY);
            yield return this.Create(nameof(this.ShadowScaleFactor), VASS.SrcConstants.PageShadowScaleFactor,
                this.ShadowScaleFactor);
            yield return this.Create(nameof(this.ShadowType), VASS.SrcConstants.PageShadowType, this.ShadowType);
            yield return this.Create(nameof(this.UIVisibility), VASS.SrcConstants.PageUIVisibility, this.UIVisibility);
            yield return this.Create(nameof(this.DrawingResizeType), VASS.SrcConstants.PageDrawingResizeType,
                this.DrawingResizeType);
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