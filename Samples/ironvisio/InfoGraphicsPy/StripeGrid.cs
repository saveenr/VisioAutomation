using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VisioAutomation.Drawing;
using BoxHierarchy=VisioAutomation.Layout.BoxHierarchy;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace InfoGraphicsPy
{
    public class StripGridItem
    {
        public  string Text;
        public  string XCategory;
        public  string YCategory;

        public StripGridItem(string text, string x, string y)
        {
            this.Text = text;
            this.XCategory = x;
            this.YCategory = y;
        }
    }

    public class RenderItem
    {
        public StripGridItem StripGridItem;
        public string Text ;
    }

    public class StripeGrid
    {
        public List<StripGridItem> Items;

        public StripeGrid()
        {
            this.Items = new List<StripGridItem>();
        }

        public StripGridItem Add(string text, string x, string y)
        {
            var item = new StripGridItem(text,x,y);
            this.Items.Add(item);
            return item;
        }

        public void Render(IVisio.Page page)
        {
            var xcats = this.Items.Select(i => i.XCategory).Distinct().ToList();
            var ycats = this.Items.Select(i => i.YCategory).Distinct().ToList();

            int cols = xcats.Count();
            int rows = ycats.Count();

            var layout = new BoxHierarchy.BoxHierarchyLayout<RenderItem>();
            layout.LayoutOptions.Origin = new VA.Drawing.Point(0,10);
            layout.LayoutOptions.DefaultHeight = 0.25;
            var root = layout.Root;
            root.Direction = BoxHierarchy.LayoutDirection.Vertical;
            root.ChildSeparation = 0.125;



            foreach (int row in Enumerable.Range(0, rows))
            {

                var n_ycat_row = root.AddNode(BoxHierarchy.LayoutDirection.Horizonal);
                //n_ycat_row.Direction = LayoutDirection.Horizonal;
                n_ycat_row.ChildSeparation = 0.25;

                foreach (int col in Enumerable.Range(0, cols))
                {
                    var n_row_col = n_ycat_row.AddNode(2.0, 0.25);

                    n_row_col.Direction = BoxHierarchy.LayoutDirection.Vertical;
                    n_row_col.AlignmentVertical = AlignmentVertical.Top;
                    var items_for_cells = this.Items.Where(i => i.XCategory == xcats[col] && i.YCategory == ycats[row]);
                    foreach (var zz in items_for_cells)
                    {
                        var n_cell = n_row_col.AddNode(2.0, 0.25);
                        var ri = new RenderItem();
                        ri.StripGridItem = zz;
                        ri.Text = zz.Text;
                        n_cell.Data = ri;
                    }
                }

                var n_ycat_title = root.AddNode(8.0, 0.5);
                var ri2 = new RenderItem();
                ri2.StripGridItem = null;
                ri2.Text = ycats[row];
                n_ycat_title.Data = ri2;
            }

            var n_xcatlabels = root.AddNode(null, 1.0);
            n_xcatlabels.Direction = BoxHierarchy.LayoutDirection.Horizonal;
            n_xcatlabels.ChildSeparation = 0.25;

            foreach (int col in Enumerable.Range(0, cols))
            {
                var n = n_xcatlabels.AddNode(2.0, 0.5);
                var ri2 = new RenderItem();
                ri2.StripGridItem = null;
                ri2.Text = xcats[col];
                n.Data = ri2;
            }

            var n_title = root.AddNode(8.0,0.5);
            var ri3 = new RenderItem();
            ri3.StripGridItem = null;
            ri3.Text = "Untitled";
            n_title.Data = ri3;
            layout.PerformLayout();

            var dom = new VA.DOM.Document();
            foreach (var n in layout.Nodes)
            {
                if (n.Data != null)
                {
                    var s = dom.DrawRectangle(n.Rectangle);
                    s.Text = n.Data.Text;
                    s.ShapeCells.VerticalAlign = 0;                    
                }
            }
            dom.Render(page);

        }
    }
}
