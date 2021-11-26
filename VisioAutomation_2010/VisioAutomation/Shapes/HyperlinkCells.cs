using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class HyperlinkCells : CellGroup
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

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.Address), Core.SrcConstants.HyperlinkAddress, this.Address);
            yield return this.Create(nameof(this.Description), Core.SrcConstants.HyperlinkDescription, this.Description);
            yield return this.Create(nameof(this.ExtraInfo), Core.SrcConstants.HyperlinkExtraInfo, this.ExtraInfo);
            yield return this.Create(nameof(this.Frame), Core.SrcConstants.HyperlinkFrame, this.Frame);
            yield return this.Create(nameof(this.SortKey), Core.SrcConstants.HyperlinkSortKey, this.SortKey);
            yield return this.Create(nameof(this.SubAddress), Core.SrcConstants.HyperlinkSubAddress, this.SubAddress);
            yield return this.Create(nameof(this.NewWindow), Core.SrcConstants.HyperlinkNewWindow, this.NewWindow);
            yield return this.Create(nameof(this.Default), Core.SrcConstants.HyperlinkDefault, this.Default);
            yield return this.Create(nameof(this.Invisible), Core.SrcConstants.HyperlinkInvisible, this.Invisible);
        }

        public static List<List<HyperlinkCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs, Core.CellValueType type)
        {
            var reader = HyperLinkCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeidpairs, type);
        }

        public static List<HyperlinkCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = HyperLinkCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<Builder> HyperLinkCells_lazy_builder = new System.Lazy<Builder>();


        class Builder : CellGroupBuilder<HyperlinkCells>
        {

            public Builder() : base(CellGroupBuilderType.MultiRow)
            {
            }

            public override HyperlinkCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new HyperlinkCells();
                var getcellvalue = row_to_cellgroup(row, cols);

                
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