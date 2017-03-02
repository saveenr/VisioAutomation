using System;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Hyperlinks
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

            if (hyperlink.Address.Formula.Value == null)
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
            hlinks_object.Address = hyperlink.Address.Formula.Value;
            hlinks_object.Description = hyperlink.Description.Formula.Value;
            hlinks_object.ExtraInfo = hyperlink.ExtraInfo.Formula.Value;
            hlinks_object.Frame= hyperlink.Frame.Formula.Value;
            hlinks_object.SubAddress= hyperlink.SubAddress.Formula.Value;
            hlinks_object.ExtraInfo= hyperlink.ExtraInfo.Formula.Value;

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

            var writer = new ShapeSheetWriter();
            hyperlink.SetFormulas(writer, row);

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
    }
}