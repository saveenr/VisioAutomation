using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class HyperlinkCells : CellRecord
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

        public override IEnumerable<CellMetadata> GetCellMetadata()
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

        public static CellRecordsGroup<HyperlinkCells> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesMultipleRows(page, shapeidpairs, type);
        }

        public static CellRecords<HyperlinkCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeMultipleRows(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        public static HyperlinkCells RowToRecord(VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols)
        {
            var record = new HyperlinkCells();
            var getcellvalue = getvalfromrowfunc(row, cols);


            record.Address = getcellvalue(nameof(Address));
            record.Description = getcellvalue(nameof(Description));
            record.ExtraInfo = getcellvalue(nameof(ExtraInfo));
            record.Frame = getcellvalue(nameof(Frame));
            record.SortKey = getcellvalue(nameof(SortKey));
            record.SubAddress = getcellvalue(nameof(SubAddress));
            record.NewWindow = getcellvalue(nameof(NewWindow));
            record.Default = getcellvalue(nameof(Default));
            record.Invisible = getcellvalue(nameof(Invisible));

            return record;
        }
        class Builder : CellRecordBuilderSectionQuery<HyperlinkCells>
        {
            public Builder() : base(HyperlinkCells.RowToRecord)
            {
            }
        }
    }
}