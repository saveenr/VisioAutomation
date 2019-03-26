using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class HyperlinkCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValueLiteral Address { get; set; }
        public VASS.CellValueLiteral Description { get; set; }
        public VASS.CellValueLiteral ExtraInfo { get; set; }
        public VASS.CellValueLiteral Frame { get; set; }
        public VASS.CellValueLiteral SortKey { get; set; }
        public VASS.CellValueLiteral SubAddress { get; set; }
        public VASS.CellValueLiteral NewWindow { get; set; }
        public VASS.CellValueLiteral Default { get; set; }
        public VASS.CellValueLiteral Invisible { get; set; }

        public override IEnumerable<VASS.CellGroups.CellMetadataItem> CellMetadata
        {
            get
            {
                yield return this.Create(nameof(this.Address), VASS.SrcConstants.HyperlinkAddress, this.Address);
                yield return this.Create(nameof(this.Description), VASS.SrcConstants.HyperlinkDescription, this.Description);
                yield return this.Create(nameof(this.ExtraInfo), VASS.SrcConstants.HyperlinkExtraInfo, this.ExtraInfo);
                yield return this.Create(nameof(this.Frame), VASS.SrcConstants.HyperlinkFrame, this.Frame);
                yield return this.Create(nameof(this.SortKey), VASS.SrcConstants.HyperlinkSortKey, this.SortKey);
                yield return this.Create(nameof(this.SubAddress), VASS.SrcConstants.HyperlinkSubAddress, this.SubAddress);
                yield return this.Create(nameof(this.NewWindow), VASS.SrcConstants.HyperlinkNewWindow, this.NewWindow);
                yield return this.Create(nameof(this.Default), VASS.SrcConstants.HyperlinkDefault, this.Default);
                yield return this.Create(nameof(this.Invisible), VASS.SrcConstants.HyperlinkInvisible, this.Invisible);
            }
        }

        public static List<List<HyperlinkCells>> GetCells(IVisio.Page page, ShapeIdPairs shapeidpairs, VASS.CellValueType type)
        {
            var reader = HyperLinkCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeidpairs, type);
        }

        public static List<HyperlinkCells> GetCells(IVisio.Shape shape, VASS.CellValueType type)
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