using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.Container
{
    public class ContainerLayout
    {
        public List<Container> Containers { get; }
        public bool IsLayedOut { get; private set; }
        public LayoutOptions LayoutOptions = new LayoutOptions();
        
        public ContainerLayout()
        {
            this.Containers = new List<Container>();
            this.IsLayedOut = false;
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
            var max_rows = this.Containers.Select(c => c.ContainerItems.Count).Max();
            var col_indexes = Enumerable.Range(0, this.Containers.Count);
            var row_indexes = Enumerable.Range(0, max_rows);

            var col_lefts =
                col_indexes.Select(i => i * (this.LayoutOptions.ItemWidth + this.LayoutOptions.ContainerHorizontalDistance + (2 * this.LayoutOptions.Padding))).ToList();

            var col_rights = col_lefts.Select(x => x + this.LayoutOptions.ItemWidth).ToList();


            var row_tops = row_indexes.Select(i => i * -(this.LayoutOptions.ItemHeight + this.LayoutOptions.ItemVerticalSpacing)).ToList();
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

                    var rect = new VisioAutomation.Geometry.Rectangle(left, bottom, right, top);

                    var item = ct.ContainerItems[ri];
                    item.Rectangle = rect;
                }
            }

            int ctn = 0;
            foreach (var ct in this.Containers)
            {
                if (ct.ContainerItems.Count < 1)
                {
                    double top = this.LayoutOptions.Padding + this.LayoutOptions.ContainerHeaderHeight;
                    double bottom = top - this.LayoutOptions.ContainerHeaderHeight - this.LayoutOptions.Padding;
                    ct.Rectangle = new VisioAutomation.Geometry.Rectangle(col_lefts[ctn], bottom, col_rights[ctn], top);
                }
                else
                {
                    double max_top = ct.ContainerItems.Select(i => i.Rectangle.Top).Max();
                    double max_right = ct.ContainerItems.Select(i => i.Rectangle.Right).Max();
                    double min_bottom = ct.ContainerItems.Select(i => i.Rectangle.Bottom).Min();
                    double min_left = ct.ContainerItems.Select(i => i.Rectangle.Left).Min();

                    max_top += this.LayoutOptions.Padding + this.LayoutOptions.ContainerHeaderHeight;
                    max_right += this.LayoutOptions.Padding;
                    min_left -= this.LayoutOptions.Padding;
                    min_bottom -= this.LayoutOptions.Padding;

                    ct.Rectangle = new VisioAutomation.Geometry.Rectangle(min_left, min_bottom, max_right, max_top);                    
                }


                ctn++;
            }

            this.IsLayedOut = true;
        }

        public IVisio.Page Render(IVisio.Document doc)
        {
            if (!this.IsLayedOut)
            {
                string msg =
                    string.Format("{0} usage error. {1}() before calling {2}().",
                        nameof(ContainerLayout), nameof(PerformLayout), nameof(Render));
                throw new System.ArgumentException(msg);
            }
            // create a new drawing
            var app = doc.Application;
            var docs = app.Documents;
            var pages = doc.Pages;
            var page = pages.Add();

            // load the stencil used to draw the items
            var item_stencil = docs.OpenStencil(this.LayoutOptions.ManualItemStencil);
            var item_stencil_masters = item_stencil.Masters;
            var item_master = item_stencil_masters[this.LayoutOptions.ManualItemMaster];
            var plain_container_master = item_stencil_masters[this.LayoutOptions.ManualContainerMaster];


            var page_shapes = page.Shapes;

            // Drop the container shapes
            var ct_items = this.Containers.ToList();
            var ct_rects = ct_items.Select(item => item.Rectangle).ToList();
            var masters = ct_items.Select(i => plain_container_master).ToList();
            short[] ct_shapeids = ContainerLayout.DropManyU(page, masters, ct_rects);

            // associate each container with the corresponding shape oject and shape id
            for (int i = 0; i < ct_items.Count; i++)
            {
                var ct_item = ct_items[i];
                var ct_shapeid = ct_shapeids[i];
                var shape = page_shapes[ct_shapeid];
                ct_item.VisioShape = shape;
                ct_item.ShapeID = ct_shapeid;
            }


            // Render the items
            var items = this.ContainerItems.ToList();
            var item_rects = items.Select(item => item.Rectangle).ToList();
            var item_masters = items.Select(i => item_master).ToList();
            short[] shapeids = ContainerLayout.DropManyU(page, item_masters, item_rects);

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

            var writer = new SidSrcWriter();

            // Format the containers and shapes

            foreach (var item in this.Containers)
            {
                this.LayoutOptions.ContainerFormatting.Apply(writer, item.ShapeID,item.ShapeID);
            }

            foreach (var item in this.ContainerItems)
            {
                this.LayoutOptions.ContainerItemFormatting.Apply(writer, item.ShapeID, item.ShapeID);
            }

            writer.BlastGuards = true;
            writer.CommitFormulas(page);

            // Set the Container Text
            foreach (var ct in this.Containers)
            {
                if (ct.Text != null)
                {
                    ct.Text.SetText(ct.VisioShape);
                }
            }

            page.ResizeToFitContents();
            app.ActiveWindow.ViewFit = (short)IVisio.VisWindowFit.visFitPage;

            return page;
        }

        private static short[] DropManyU(
            IVisio.Page page,
            IList<IVisio.Master> masters,
            IList<VisioAutomation.Geometry.Rectangle> rects)
        {
            var points = rects.Select(r => r.Center).ToList();
            var shapeids = page.DropManyU(masters, points);

            // Dropping takes care of the PinX and PinY
            // Now set the Width's and Heights
            var writer = new SidSrcWriter();
            for (int i = 0; i < rects.Count; i++)
            {
                writer.SetValue(shapeids[i], VisioAutomation.ShapeSheet.SrcConstants.XFormWidth, rects[i].Width);
                writer.SetValue(shapeids[i], VisioAutomation.ShapeSheet.SrcConstants.XFormHeight, rects[i].Height);
            }

            writer.CommitFormulas(page);

            return shapeids;
        }
    }

}
