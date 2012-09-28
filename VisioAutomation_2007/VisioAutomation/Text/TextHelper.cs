using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Text
{
    public static class TextHelper
    {
        public static void FitShapeToText(IVisio.Page page, IEnumerable<IVisio.Shape> shapes)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            if (shapes == null)
            {
                throw new System.ArgumentNullException("shapes");
            }

            var shapeids = shapes.Select(s => s.ID).ToList();

            // Calculate the new sizes for each shape
            var new_sizes = new List<VA.Drawing.Size>(shapeids.Count);
            foreach (var shape in shapes)
            {
                var text_bounding_box = shape.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightText).Size;
                var wh_bounding_box = shape.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightWH).Size;

                double max_w = System.Math.Max(text_bounding_box.Width, wh_bounding_box.Width);
                double max_h = System.Math.Max(text_bounding_box.Height, wh_bounding_box.Height);
                var max_size = new VA.Drawing.Size(max_w, max_h);
                new_sizes.Add(max_size);
            }

            var src_width = VA.ShapeSheet.SRCConstants.Width;
            var src_height = VA.ShapeSheet.SRCConstants.Height;

            var update = new VA.ShapeSheet.Update();
            for (int i = 0; i < new_sizes.Count; i++)
            {
                var shapeid = shapeids[i];
                var new_size = new_sizes[i];
                update.SetFormula((short) shapeid, src_width, new_size.Width);
                update.SetFormula((short) shapeid, src_height, new_size.Height);
            }

            update.Execute(page);
        }


        public static IVisio.Font TryGetFont(IVisio.Fonts fonts, string name)
        {
            try
            {
                var font = fonts[name];
                return font;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }
    }
}