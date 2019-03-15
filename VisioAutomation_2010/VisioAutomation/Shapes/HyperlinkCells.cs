using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;

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
    }
}