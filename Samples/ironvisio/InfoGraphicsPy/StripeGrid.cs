using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VisioAutomation.Layout.BoxHierarchy;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace InfoGraphicsPy
{
    public class StripGridItem
    {
        public  string Text;
        public  string XCategory;
        public  string YCategory;
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
            var item = new StripGridItem();
            item.Text = text;
            item.XCategory = x;
            item.YCategory = y;

            this.Items.Add(item);
            return item;
        }

        public void Render(IVisio.Page page)
        {
            var xcats = this.Items.Select(i => i.XCategory).Distinct().ToList();
            var ycats = this.Items.Select(i => i.YCategory).Distinct().ToList();

            int cols = xcats.Count();
            int rows = ycats.Count();

            double origin_y = 8.0;
            double boty = 0;

            var layout = new VA.Layout.BoxHierarchy.BoxHierarchyLayout<string>();
            layout.LayoutOptions.Origin = new VA.Drawing.Point(0,10);
            var root = layout.Root;
            root.Direction = LayoutDirection.Vertical;
            var n_toprow = root.AddNode(null, 1.0);
            n_toprow.Direction = LayoutDirection.Horizonal;
            n_toprow.ChildSeparation = 0.25;

            // draw xcat bars
            foreach (int col in Enumerable.Range(0, cols))
            {
                var n = n_toprow.AddNode(2.0, 0.5);
                n.Data = xcats[col];
            }

            var n_toprow1 = root.AddNode(null, 2.0);

            foreach (int row in Enumerable.Range(0, rows))
            {
                var n_ycat_title = root.AddNode(8.0, 1.0);
                n_ycat_title.Data = ycats[row];

                var n_ycat_row = root.AddNode(8.0, 1.0);
                n_ycat_row.Direction = LayoutDirection.Horizonal;
                n_ycat_row.ChildSeparation = 0.25;
                foreach (int col in Enumerable.Range(0, cols))
                {
                    var n_row_col = n_ycat_row.AddNode(2.0, 0.5);
                    n_row_col.Data = "";

                    n_row_col.Direction = LayoutDirection.Vertical;
                    var z = this.Items.Where(i => i.XCategory == xcats[col] && i.YCategory == ycats[row]);
                    foreach (var zz in z)
                    {
                        var n_cell = n_row_col.AddNode(2.0, 0.5);
                        n_cell.Data = zz.Text;
                    }
                }

            }

            layout.PerformLayout();

            var dom = new VA.DOM.Document();

            var ts = dom.DrawRectangle(n_toprow1.Rectangle);
            ts.Text = "Title";


            foreach (var n in layout.Nodes)
            {
                if (n.Data != null)
                {
                    var s = dom.DrawRectangle(n.Rectangle);
                    s.Text = n.Data;
                    s.ShapeCells.VerticalAlign = 0;                    
                }
            }
           

            dom.Render(page);

        }
    }
}
