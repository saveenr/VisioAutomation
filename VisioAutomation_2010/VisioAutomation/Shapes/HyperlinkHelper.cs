using IVisio = Microsoft.Office.Interop.Visio;

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
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (hyperlink == null)
            {
                throw new System.ArgumentNullException(nameof(hyperlink));
            }

            if (hyperlink.Address.Value == null)
            {
                throw new System.ArgumentException("Address is null", nameof(hyperlink));
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
            hlinks_object.Frame = hyperlink.Frame.Value;
            hlinks_object.SubAddress = hyperlink.SubAddress.Value;
            hlinks_object.ExtraInfo = hyperlink.ExtraInfo.Value;

            // The Hyperlink object has no property for these cells, so write them to the shape sheet.
            // Cells that were not set are skipped by the writer.
            var other_cells = new HyperlinkCells();
            if (hyperlink.SortKey.HasValue)
            {
                // the sort key is a string cell, so it needs quoting to be a valid formula
                other_cells.SortKey = Core.CellValue.EncodeValue(hyperlink.SortKey.Value);
            }
            other_cells.NewWindow = hyperlink.NewWindow;
            other_cells.Default = hyperlink.Default;
            other_cells.Invisible = hyperlink.Invisible;

            short row = hlinks_object.Row;
            var writer = new ShapeSheet.Writers.SrcWriter();
            writer.SetValues(other_cells, row);
            writer.Commit(shape, Core.CellValueType.Formula);

            return row;
        }

        public static int Set(
            IVisio.Shape shape,
            short row,
            HyperlinkCells hyperlink)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var writer = new ShapeSheet.Writers.SrcWriter();
            writer.SetValues(hyperlink, row);

            writer.Commit(shape, Core.CellValueType.Formula);

            return row;
        }

        public static void Delete(IVisio.Shape shape, int index)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (index < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(index));
            }

            var row = (IVisio.VisRowIndices) index;
            shape.DeleteRow((short) IVisio.VisSectionIndices.visSectionHyperlink, (short) row);
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            return shape.RowCount[(short) IVisio.VisSectionIndices.visSectionHyperlink];
        }
    }
}