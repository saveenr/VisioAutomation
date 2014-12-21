using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Models.ContainerLayout
{
    public class ContainerLayout
    {
        public List<Container> Containers { get; private set; }
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

                    var rect = new VA.Drawing.Rectangle(left, bottom, right, top);

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
                    ct.Rectangle = new VA.Drawing.Rectangle(col_lefts[ctn], bottom, col_rights[ctn], top);
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

                    ct.Rectangle = new VA.Drawing.Rectangle(min_left, min_bottom, max_right, max_top);                    
                }


                ctn++;
            }

            this.IsLayedOut = true;
        }

        public IVisio.Page Render(IVisio.Document doc)
        {
            if (!this.IsLayedOut)
            {
                string msg = string.Format("{0} usage error. Call PerformLayout() before calling Render().",
                                           typeof (ContainerLayout).Name);
                throw new VA.AutomationException(msg);
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

            var update = new VA.ShapeSheet.Update();

            // Format the containers and shapes

            foreach (var item in this.Containers)
            {
                this.LayoutOptions.ContainerFormatting.Apply(update, item.ShapeID,item.ShapeID);
            }

            foreach (var item in this.ContainerItems)
            {
                this.LayoutOptions.ContainerItemFormatting.Apply(update, item.ShapeID, item.ShapeID);
            }     

            update.BlastGuards = true;
            update.Execute(page);

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
            IList<VA.Drawing.Rectangle> rects)
        {
            var points = rects.Select(r => r.Center).ToList();
            var shapeids = VA.Pages.PageHelper.DropManyU(page, masters, points);

            var xfrm = new VA.Shapes.XFormCells();

            var update = new VA.ShapeSheet.Update(points.Count*2);
            for (int i = 0; i < rects.Count(); i++)
            {
                xfrm.Width = rects[i].Width;
                xfrm.Height = rects[i].Height;
                update.SetFormulas(shapeids[i], xfrm);
            }
            update.Execute(page);

            return shapeids;
        }
    }

}
