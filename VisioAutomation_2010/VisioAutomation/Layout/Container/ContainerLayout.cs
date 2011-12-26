using System.Collections.Generic;
using System.Linq;
using System.Text;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.ContainerLayout
{
    public class ContainerLayout
    {
        public List<Container> Containers { get; private set; }
        public LayoutOptions LayoutOptions = new LayoutOptions();

        public ContainerLayout()
        {
            this.Containers = new List<Container>();
        }

        public Container AddContainer(string text)
        {
            var c = new Container(text);
            this.Containers.Add(c);
            return c;
        }

        public IEnumerable<ContainerItem> ContainerItems
        {
            get
            {
                foreach (var c in this.Containers)
                {
                    foreach (var item in c.ContainerItems)
                    {
                        yield return item;
                    }
                }
            }
        }

        public void PerformLayout()
        {
            var origin = new VA.Drawing.Point(0, 0);


            var col_lefts =
                Enumerable.Range(0, this.Containers.Count).Select(i => i * (this.LayoutOptions.ItemWidth + this.LayoutOptions.ContainerHorizontalDistance + (2 * this.LayoutOptions.Padding))).ToList();

            var col_rights = col_lefts.Select(x => x + this.LayoutOptions.ItemWidth).ToList();

            var max_rows = this.Containers.Select(c => c.ContainerItems.Count).Max();

            var row_tops = Enumerable.Range(0, max_rows).Select(i => i * -(this.LayoutOptions.ItemHeight + this.LayoutOptions.ItemVerticalSpacing)).ToList();
            var row_bottoms = row_tops.Select(y => y - this.LayoutOptions.ItemHeight).ToList();

            for (int container = 0; container< this.Containers.Count; container++)
            {
                var ct = this.Containers[container];
                for (int ri=0;ri<ct.ContainerItems.Count;ri++)
                {
                    double left = col_lefts[container];
                    double right = col_rights[container];
                    double top = row_tops[ri];
                    double bottom = row_bottoms[ri];

                    var rect = new VA.Drawing.Rectangle(left, bottom, right, top);

                    var item = ct.ContainerItems[ri];
                    item.Rectangle = rect;
                }
            }

            // Calculate a rectangle for the container for rendering that doesn't use the container API
            foreach (var ct in this.Containers)
            {
                double max_top = ct.ContainerItems.Select(i => i.Rectangle.Top).Max();
                double max_right = ct.ContainerItems.Select(i => i.Rectangle.Right).Max();
                double min_bottom = ct.ContainerItems.Select(i => i.Rectangle.Bottom).Min();
                double min_left = ct.ContainerItems.Select(i => i.Rectangle.Left).Min();

                max_top += this.LayoutOptions.Padding + this.LayoutOptions.ContainerHeaderHeight;
                max_right += this.LayoutOptions.Padding;
                min_left -= this.LayoutOptions.Padding;
                min_bottom -= this.LayoutOptions.Padding;
                
                ct.Rectangle = new VA.Drawing.Rectangle(min_left, min_bottom, max_right, max_top);
            }
        }

        public void Render(IVisio.Application app)
        {
            var docs = app.Documents;
            var doc = docs.Add("");

            IVisio.Master special_container_master=null;

            if (this.LayoutOptions.Style == RenderStyle.UseVisioContainers)
            {
                // only load the special Container stencil if needed.
                
                doc.DiagramServicesEnabled = (int)IVisio.VisDiagramServices.visServiceVersion140;
                // load the special container stencil
                var measurement = IVisio.VisMeasurementSystem.visMSUS;
                var stenciltype = IVisio.VisBuiltInStencilTypes.visBuiltInStencilContainers;
                string stencilfile = app.GetBuiltInStencilFile(stenciltype, measurement);
                short flags = (short)IVisio.VisOpenSaveArgs.visAddDocked;
                var container_stencil = docs.OpenEx(stencilfile, flags);

                var container_stencil_masters = container_stencil.Masters;
                special_container_master = container_stencil_masters["Container 1"];               
            }

            // load the stencil used to draw the items
            var item_stencil = VA.DocumentHelper.OpenStencil(docs, "basic_u.vss");
            var item_stencil_masters = item_stencil.Masters;
            var item_master = item_stencil_masters["Rounded Rectangle"];
            var plain_container_master = item_stencil_masters["Rectangle"];

            // create a new drawing
            var page = doc.Pages[1];

            var page_shapes = page.Shapes;

            // Render containers withou using container API
            if (this.LayoutOptions.Style == VA.Layout.ContainerLayout.RenderStyle.UseShapes)
            {
                // Drop the container shapes
                var ct_items = this.Containers.ToList();
                var ct_rects = ct_items.Select(item => item.Rectangle).ToList();
                var masters = ct_items.Select(i => plain_container_master).ToList();
                short[] ct_shapeids = DropManyU(page, masters, ct_rects);

                // associate each container with the corresponding shape oject and shape id
                for (int i = 0; i < ct_items.Count; i++)
                {
                    var ct_item = ct_items[i];
                    var ct_shapeid = ct_shapeids[i];
                    var shape = page_shapes[ct_shapeid];
                    ct_item.VisioShape = shape;
                    ct_item.ShapeID = ct_shapeid;
                }

            }

            // Render the items
            var items = this.ContainerItems.ToList();
            var item_rects = items.Select(item => item.Rectangle).ToList();
            var item_masters = items.Select(i => item_master).ToList();
            short[] shapeids = DropManyU(page, item_masters, item_rects);

            // Associate each item with the corresponding shape object and shape id
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var shapeid = shapeids[i];
                var shape = page_shapes[shapeid];
                item.VisioShape = shape;
                item.ShapeID = shapeid;
            }

            // Often useful to show everthing because these diagrams can get large
            app.ActiveWindow.ViewFit = (short)IVisio.VisWindowFit.visFitPage;

            // Set the items
            foreach (var item in items.Where(i => i.Text != null))
            {
                item.VisioShape.Text = item.Text;
            }

            var window = app.ActiveWindow;

            // Render containers using container API
            if (this.LayoutOptions.Style == VA.Layout.ContainerLayout.RenderStyle.UseVisioContainers)
            {
                foreach (var ct in this.Containers)
                {
                    window.DeselectAll();
                    foreach (var item in ct.ContainerItems)
                    {
                        window.Select(item.VisioShape, (short)IVisio.VisSelectArgs.visSelect);
                    }
                    var sel = window.Selection;

                    ct.VisioShape = page.DropContainer(special_container_master, sel);
                    ct.ShapeID = ct.VisioShape.ID16;
                }                
            }

            // Set the Container Text
            foreach (var ct in this.Containers)
            {
                ct.VisioShape.Text = ct.Text;
            }

            // Format the containers and shapes
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            foreach (var item in this.ContainerItems)
            {
                if (item.CharacterFormatCells != null)
                {
                    item.CharacterFormatCells.Apply(update, item.ShapeID, 0);
                }
                if (item.ParagraphFormatCells != null)
                {
                    item.ParagraphFormatCells.Apply(update, item.ShapeID, 0);
                }
                if (item.ShapeFormatCells != null)
                {
                    item.ShapeFormatCells.Apply(update, item.ShapeID);
                }
                if (item.TextBlockFormatCells != null)
                {
                    item.TextBlockFormatCells.Apply(update, item.ShapeID);
                }
            }

            // Unless we do this application of these properties will fail
            // because some cells are guarded by default
            update.BlastGuards = true;

            update.Execute(page);

            page.ResizeToFitContents();
            app.ActiveWindow.ViewFit = (short)IVisio.VisWindowFit.visFitPage;
        }

        private static short[] DropManyU(
            IVisio.Page page,
            IList<IVisio.Master> masters,
            IList<VA.Drawing.Rectangle> rects)
        {
            var points = rects.Select(r => r.Center).ToList();
            var shapeids = VA.Pages.PageHelper.DropManyU(page, masters, points);

            var xfrm = new VA.Layout.XFormCells();

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate(points.Count*2);
            for (int i = 0; i < rects.Count(); i++)
            {
                xfrm.Width = rects[i].Width;
                xfrm.Height = rects[i].Height;
                xfrm.Apply(update,shapeids[i]);
            }
            update.Execute(page);

            return shapeids;
        }
    }

}
