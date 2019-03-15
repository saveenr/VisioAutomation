using System;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public static class HyperlinkHelper
    {

        public static int Add(
            IVisio.Shape shape,
            HyperlinkCells hyperlink)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            if (hyperlink == null)
            {
                throw new ArgumentNullException(nameof(hyperlink));
            }

            if (hyperlink.Address.Value == null)
            {
                throw new ArgumentException("Address is null",nameof(hyperlink));
            }

            /*
            TODO: Why doesn't this work?
            short row = shape.AddRow((short)IVisio.VisSectionIndices.visSectionHyperlink,
                                     (short)IVisio.VisRowIndices.visRowLast,
                                     (short)IVisio.VisRowTags.visTagDefault);

            HyperlinkHelper.Set(shape, row, hyperlink);

    */
            var hlinks_collection = shape.Hyperlinks;
            var hlinks_object = hlinks_collection.Add();
            hlinks_object.Address = hyperlink.Address.Value;
            hlinks_object.Description = hyperlink.Description.Value;
            hlinks_object.ExtraInfo = hyperlink.ExtraInfo.Value;
            hlinks_object.Frame= hyperlink.Frame.Value;
            hlinks_object.SubAddress= hyperlink.SubAddress.Value;
            hlinks_object.ExtraInfo= hyperlink.ExtraInfo.Value;

            //hlinks_object.NewWindow = hyperlink.NewWindow.Formula.Value;
            //hlinks_object.IsDefaultLink = hyperlink.Default.Formula.Value;
            // hlinks_object.XXX = hyperlink.Invisible.Formula.Value;

            return hlinks_object.Row;
        }

        public static int Set(
            IVisio.Shape shape,
            short row,
            HyperlinkCells hyperlink)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetFormulas(hyperlink, row);

            writer.Commit(shape);

            return row;
        }

        public static void Delete(IVisio.Shape shape, int index)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var row = (IVisio.VisRowIndices)index;
            shape.DeleteRow( (short) IVisio.VisSectionIndices.visSectionHyperlink, (short)row);
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            return shape.RowCount[(short)IVisio.VisSectionIndices.visSectionHyperlink];
        }

        public static List<List<HyperlinkCells>> GetHyperlinkCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = HyperLinkCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<HyperlinkCells> GetHyperlinkCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = HyperLinkCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<HyperlinkCellsReader> HyperLinkCells_lazy_reader = new System.Lazy<HyperlinkCellsReader>();


        class HyperlinkCellsReader : CellGroupReader<HyperlinkCells>
        {

            public HyperlinkCellsReader() : base(new VisioAutomation.ShapeSheet.Query.SectionsQuery())
            {

                var sec = this.query_multirow.SectionQueries.Add(IVisio.VisSectionIndices.visSectionHyperlink);
                var temp_cells = new HyperlinkCells();
                foreach (var pair in temp_cells.CellMetadata)
                {
                    sec.Columns.Add(pair.Src, pair.Name);
                }

            }

            public override HyperlinkCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new HyperlinkCells();

                var cols = this.query_multirow.SectionQueries[0].Columns;

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