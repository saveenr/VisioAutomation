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


            var dom = new VA.DOM.Document();

            
            // draw xcat bars
            foreach (int col in Enumerable.Range(0,cols))
            {
                double left_ = 0 + ((1.0 + 0.25)*col);
                double right = left_+1.0;
                double top = origin_y;
                double bottom = boty;

                var r = dom.DrawRectangle(left_, bottom, right, top);
                r.Text = xcats[col];
                r.ShapeCells.VerticalAlign = 0;
            }

            dom.Render(page);

        }
    }
}
