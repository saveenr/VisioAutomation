using System.Collections.Generic;
using System.Linq;
using System.Text;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

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
            var container_model = this;
            var docs = app.Documents;

            // load the special container stencil
            var container_stencil = Layout.ContainerLayout.ContainerUtil.LoadContainerStencil(docs);
            var container_stencil_masters = container_stencil.Masters;
            var container_master = container_stencil_masters["Container 1"];

            // load the special container stencil
            var basic_stencil = VA.DocumentHelper.OpenStencil(docs, "basic_u.vss");
            var basic_stencil_masters = basic_stencil.Masters;
            var basic_master = basic_stencil_masters["Rounded Rectangle"];


            // create a new drawing
            var doc = docs.Add("");
            var page = doc.Pages[1];

            container_model.PerformLayout();



            // Render containers withou using container API
            if (this.LayoutOptions.RenderWithShapes == true )
            {
                var ct_items = container_model.Containers.ToList();
                var ct_rects = ct_items.Select(item => item.Rectangle).ToList();
                var masters = ct_items.Select(i => basic_master).ToList();
                short[] ct_shapeids = DropManyU(page, masters, ct_rects);
                var xpage_shapes = page.Shapes;

                for (int i = 0; i < ct_items.Count; i++)
                {
                    var ct_item = ct_items[i];
                    var ct_shapeid = ct_shapeids[i];
                    var shape = xpage_shapes[ct_shapeid];
                    ct_item.VisioShape = shape;
                    ct_item.ShapeID = ct_shapeid;
                }
            }

            var items = container_model.ContainerItems.ToList();
            var rects = items.Select(item => item.Rectangle).ToList();
            var masters2 = items.Select(i => basic_master).ToList();
            short[] shapeids = DropManyU(page, masters2, rects);
            var page_shapes = page.Shapes;

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var shapeid = shapeids[i];
                var shape = page_shapes[shapeid];
                item.VisioShape = shape;
                item.ShapeID = shapeid;
            }

            foreach (var item in items.Where(i => i.Text != null))
            {
                item.VisioShape.Text = item.Text;
            }

            var window = app.ActiveWindow;

            // Render containers using container API
            if (this.LayoutOptions.RenderWithShapes==false)
            {
                foreach (var ct in container_model.Containers)
                {

                    window.DeselectAll();
                    foreach (var item in ct.ContainerItems)
                    {
                        window.Select(item.VisioShape, (short)IVisio.VisSelectArgs.visSelect);
                    }
                    var sel = window.Selection;

                    ct.VisioShape = page.DropContainer(container_master, sel);
                    ct.ShapeID = ct.VisioShape.ID16;
                }                
            }


            // Set the Container Text
            foreach (var ct in container_model.Containers)
            {
                ct.VisioShape.Text = ct.Text;
            }
            
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            foreach (var ct in container_model.ContainerItems)
            {
                if (ct.CharacterFormatCells != null)
                {
                    ct.CharacterFormatCells.Apply(update, ct.ShapeID, 0);
                }
                if (ct.ParagraphFormatCells != null)
                {
                    ct.ParagraphFormatCells.Apply(update, ct.ShapeID, 0);
                }
                if (ct.ShapeFormatCells != null)
                {
                    ct.ShapeFormatCells.Apply(update, ct.ShapeID);
                }
                if (ct.TextBlockFormatCells != null)
                {
                    ct.TextBlockFormatCells.Apply(update, ct.ShapeID);
                }
            }

            // Unless we do this application of these properties will fail
            // because some cells are guarded by default
            update.BlastGuards = true;

            update.Execute(page);

            page.ResizeToFitContents();
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

                xfrm.Apply(update, shapeids[i]);
            }

            update.Execute(page);

            return shapeids;
        }
    }

}
