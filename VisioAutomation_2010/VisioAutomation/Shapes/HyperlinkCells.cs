using System.Collections.Generic;
using VACG = VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class HyperlinkCells : VACG.CellGroup
    {
        public Core.CellValue Address { get; set; }
        public Core.CellValue Description { get; set; }
        public Core.CellValue ExtraInfo { get; set; }
        public Core.CellValue Frame { get; set; }
        public Core.CellValue SortKey { get; set; }
        public Core.CellValue SubAddress { get; set; }
        public Core.CellValue NewWindow { get; set; }
        public Core.CellValue Default { get; set; }
        public Core.CellValue Invisible { get; set; }

        public override IEnumerable<VACG.CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.Address), Core.SrcConstants.HyperlinkAddress, this.Address);
            yield return this._create(nameof(this.Description), Core.SrcConstants.HyperlinkDescription,
                this.Description);
            yield return this._create(nameof(this.ExtraInfo), Core.SrcConstants.HyperlinkExtraInfo, this.ExtraInfo);
            yield return this._create(nameof(this.Frame), Core.SrcConstants.HyperlinkFrame, this.Frame);
            yield return this._create(nameof(this.SortKey), Core.SrcConstants.HyperlinkSortKey, this.SortKey);
            yield return this._create(nameof(this.SubAddress), Core.SrcConstants.HyperlinkSubAddress, this.SubAddress);
            yield return this._create(nameof(this.NewWindow), Core.SrcConstants.HyperlinkNewWindow, this.NewWindow);
            yield return this._create(nameof(this.Default), Core.SrcConstants.HyperlinkDefault, this.Default);
            yield return this._create(nameof(this.Invisible), Core.SrcConstants.HyperlinkInvisible, this.Invisible);
        }

        public static List<List<HyperlinkCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesMultipleRows(page, shapeidpairs, type);
        }

        public static List<HyperlinkCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeMultipleRows(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();


        class Builder : VACG.CellGroupBuilder<HyperlinkCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.MultiRow)
            {
            }

            public override HyperlinkCells ToCellGroup(VASS.Data.CellValueRow<string> row, VASS.Query.Columns cols)
            {
                var cells = new HyperlinkCells();
                var getcellvalue = queryrow_to_cellgroup(row, cols);


                cells.Address = getcellvalue(nameof(Address));
                cells.Description = getcellvalue(nameof(Description));
                cells.ExtraInfo = getcellvalue(nameof(ExtraInfo));
                cells.Frame = getcellvalue(nameof(Frame));
                cells.SortKey = getcellvalue(nameof(SortKey));
                cells.SubAddress = getcellvalue(nameof(SubAddress));
                cells.NewWindow = getcellvalue(nameof(NewWindow));
                cells.Default = getcellvalue(nameof(Default));
                cells.Invisible = getcellvalue(nameof(Invisible));

                return cells;
            }
        }
    }
}