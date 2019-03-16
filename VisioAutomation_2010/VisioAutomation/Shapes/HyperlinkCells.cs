using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class HyperlinkCells : CellGroup
    {
        public CellValueLiteral Address { get; set; }
        public CellValueLiteral Description { get; set; }
        public CellValueLiteral ExtraInfo { get; set; }
        public CellValueLiteral Frame { get; set; }
        public CellValueLiteral SortKey { get; set; }
        public CellValueLiteral SubAddress { get; set; }
        public CellValueLiteral NewWindow { get; set; }
        public CellValueLiteral Default { get; set; }
        public CellValueLiteral Invisible { get; set; }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {
                yield return CellMetadataItem.Create(nameof(this.Address), SrcConstants.HyperlinkAddress, this.Address);
                yield return CellMetadataItem.Create(nameof(this.Description), SrcConstants.HyperlinkDescription, this.Description);
                yield return CellMetadataItem.Create(nameof(this.ExtraInfo), SrcConstants.HyperlinkExtraInfo, this.ExtraInfo);
                yield return CellMetadataItem.Create(nameof(this.Frame), SrcConstants.HyperlinkFrame, this.Frame);
                yield return CellMetadataItem.Create(nameof(this.SortKey), SrcConstants.HyperlinkSortKey, this.SortKey);
                yield return CellMetadataItem.Create(nameof(this.SubAddress), SrcConstants.HyperlinkSubAddress, this.SubAddress);
                yield return CellMetadataItem.Create(nameof(this.NewWindow), SrcConstants.HyperlinkNewWindow, this.NewWindow);
                yield return CellMetadataItem.Create(nameof(this.Default), SrcConstants.HyperlinkDefault, this.Default);
                yield return CellMetadataItem.Create(nameof(this.Invisible), SrcConstants.HyperlinkInvisible, this.Invisible);
            }
        }

        public static List<List<HyperlinkCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = HyperLinkCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<HyperlinkCells> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = HyperLinkCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<HyperlinkCellsBuilder> HyperLinkCells_lazy_builder = new System.Lazy<HyperlinkCellsBuilder>();


        class HyperlinkCellsBuilder : CellGroupBuilder<HyperlinkCells>
        {

            public HyperlinkCellsBuilder() : base(CellGroupBuilderType.MultiRow)
            {
            }

            public override HyperlinkCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new HyperlinkCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }


                cells.Address = getcellvalue(nameof(HyperlinkCells.Address));
                cells.Description = getcellvalue(nameof(HyperlinkCells.Description));
                cells.ExtraInfo = getcellvalue(nameof(HyperlinkCells.ExtraInfo));
                cells.Frame = getcellvalue(nameof(HyperlinkCells.Frame));
                cells.SortKey = getcellvalue(nameof(HyperlinkCells.SortKey));
                cells.SubAddress = getcellvalue(nameof(HyperlinkCells.SubAddress));
                cells.NewWindow = getcellvalue(nameof(HyperlinkCells.NewWindow));
                cells.Default = getcellvalue(nameof(HyperlinkCells.Default));
                cells.Invisible = getcellvalue(nameof(HyperlinkCells.Invisible));

                return cells;
            }
        }


    }
}