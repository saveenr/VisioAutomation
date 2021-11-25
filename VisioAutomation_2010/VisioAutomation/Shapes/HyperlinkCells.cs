using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class HyperlinkCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue Address { get; set; }
        public VisioAutomation.Core.CellValue Description { get; set; }
        public VisioAutomation.Core.CellValue ExtraInfo { get; set; }
        public VisioAutomation.Core.CellValue Frame { get; set; }
        public VisioAutomation.Core.CellValue SortKey { get; set; }
        public VisioAutomation.Core.CellValue SubAddress { get; set; }
        public VisioAutomation.Core.CellValue NewWindow { get; set; }
        public VisioAutomation.Core.CellValue Default { get; set; }
        public VisioAutomation.Core.CellValue Invisible { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.Address), VisioAutomation.Core.SrcConstants.HyperlinkAddress, this.Address);
            yield return this.Create(nameof(this.Description), VisioAutomation.Core.SrcConstants.HyperlinkDescription, this.Description);
            yield return this.Create(nameof(this.ExtraInfo), VisioAutomation.Core.SrcConstants.HyperlinkExtraInfo, this.ExtraInfo);
            yield return this.Create(nameof(this.Frame), VisioAutomation.Core.SrcConstants.HyperlinkFrame, this.Frame);
            yield return this.Create(nameof(this.SortKey), VisioAutomation.Core.SrcConstants.HyperlinkSortKey, this.SortKey);
            yield return this.Create(nameof(this.SubAddress), VisioAutomation.Core.SrcConstants.HyperlinkSubAddress, this.SubAddress);
            yield return this.Create(nameof(this.NewWindow), VisioAutomation.Core.SrcConstants.HyperlinkNewWindow, this.NewWindow);
            yield return this.Create(nameof(this.Default), VisioAutomation.Core.SrcConstants.HyperlinkDefault, this.Default);
            yield return this.Create(nameof(this.Invisible), VisioAutomation.Core.SrcConstants.HyperlinkInvisible, this.Invisible);
        }

        public static List<List<HyperlinkCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs, VisioAutomation.Core.CellValueType type)
        {
            var reader = HyperLinkCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeidpairs, type);
        }

        public static List<HyperlinkCells> GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = HyperLinkCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<HyperlinkCellsBuilder> HyperLinkCells_lazy_builder = new System.Lazy<HyperlinkCellsBuilder>();


        class HyperlinkCellsBuilder : VASS.CellGroups.CellGroupBuilder<HyperlinkCells>
        {

            public HyperlinkCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.MultiRow)
            {
            }

            public override HyperlinkCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new HyperlinkCells();
                var getcellvalue = VASS.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                
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