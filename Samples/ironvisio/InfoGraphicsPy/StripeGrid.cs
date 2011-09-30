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
        public string Text;
        public string XCategory;
        public string YCategory;

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
        public VA.DOM.ShapeCells ShapeCells;
        public bool Underline;
    }

    public class StripeGrid
    {
        public List<StripGridItem> Items;
        public bool ToUpper;
        public string Title = "Untitled";

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
            var titleformat = new VA.DOM.ShapeCells();
            titleformat.VerticalAlign = 0;
            titleformat.HAlign = 0;
            titleformat.CharSize = VA.Convert.PointsToInches(24);
            titleformat.LinePattern = 0;
            titleformat.LineWeight = 0;

            var cellformat = new VA.DOM.ShapeCells();
            cellformat.VerticalAlign = 0;
            cellformat.HAlign = 0;
            cellformat.CharSize = VA.Convert.PointsToInches(8);
            cellformat.LinePattern = 0;
            cellformat.LineWeight = 0;
            cellformat.FillForegnd = "rgb(240,240,240)";

            var xcatformat = new VA.DOM.ShapeCells();
            xcatformat.VerticalAlign = 0;
            xcatformat.HAlign = 0;
            xcatformat.CharSize = VA.Convert.PointsToInches(14);
            xcatformat.LinePattern = 0;
            xcatformat.LineWeight = 0;
            xcatformat.CharStyle = ((int)VA.Text.CharStyle.Bold);

            var ycatformat = new VA.DOM.ShapeCells();
            ycatformat.VerticalAlign = 2;
            ycatformat.HAlign = 0;
            ycatformat.CharSize = VA.Convert.PointsToInches(14);
            ycatformat.LinePattern = 0;
            ycatformat.LineWeight = 0;
            //ycatformat.FillForegnd = "rgb(220,230,255)";
            ycatformat.CharStyle = ((int)VA.Text.CharStyle.Bold);

            var xcats = this.Items.Select(i => i.XCategory).Distinct().ToList();
            var ycats = this.Items.Select(i => i.YCategory).Distinct().ToList();

            int cols = xcats.Count();
            int rows = ycats.Count();

            double cell_width = 3.0;
            double cell_vsep = 0.125;
            double cell_height = 0.25;
            double indent = 2.0;
            double pwidth = cols * cell_width + (System.Math.Max(0, cols - 1) * 0.25) + indent + 0.25;
            double YCatTitleHeight = 0.5;
            double colsep = 0.25;

            var layout = new BoxHierarchy.BoxHierarchyLayout<RenderItem>();
            layout.LayoutOptions.Origin = new VA.Drawing.Point(0,10);
            layout.LayoutOptions.DefaultHeight = 0.25;
            var root = layout.Root;
            root.Direction = BoxHierarchy.LayoutDirection.Vertical;
            root.ChildSeparation = 0.125;

            foreach (int row in Enumerable.Range(0, rows))
            {

                var n_ycat_row = root.AddNode(BoxHierarchy.LayoutDirection.Horizonal);
                n_ycat_row.ChildSeparation = colsep;

                // -- add indent
                n_ycat_row.AddNode(indent, 0.25);

                foreach (int col in Enumerable.Range(0, cols))
                {
                    var n_row_col = n_ycat_row.AddNode(cell_width, 0.25);

                    // ---
                    n_row_col.Direction = BoxHierarchy.LayoutDirection.Vertical;
                    n_row_col.AlignmentVertical = AlignmentVertical.Top;
                    n_row_col.ChildSeparation = cell_vsep;
                    var items_for_cells = this.Items.Where(i => i.XCategory == xcats[col] && i.YCategory == ycats[row]);
                    foreach (var cell_item in items_for_cells)
                    {
                        var n_cell = n_row_col.AddNode(cell_width, cell_height);
                        var cell_data = new RenderItem();
                        cell_data.StripGridItem = cell_item;
                        cell_data.Text = cell_item.Text;
                        cell_data.ShapeCells = cellformat;
                        n_cell.Data = cell_data;
                    }
                }

                var n_ycat_title = root.AddNode(pwidth, YCatTitleHeight);
                var ycat_data = new RenderItem();
                ycat_data.StripGridItem = null;
                ycat_data.Text = ycats[row];
                ycat_data.ShapeCells = ycatformat;
                ycat_data.Underline = true;
                n_ycat_title.Data = ycat_data;
            }

            var n_xcatlabels = root.AddNode(null, 1.0);
            n_xcatlabels.Direction = BoxHierarchy.LayoutDirection.Horizonal;
            n_xcatlabels.ChildSeparation = colsep;

            // Add indent
            n_xcatlabels.AddNode(indent, 0.25);

            // Add XCategory labels
            foreach (int col in Enumerable.Range(0, cols))
            {
                var n = n_xcatlabels.AddNode(cell_width, 0.5);
                var xcat_data = new RenderItem();
                xcat_data.StripGridItem = null;
                xcat_data.Text = xcats[col];
                xcat_data.ShapeCells = xcatformat;
                n.Data = xcat_data;
            }

            // Add Title for Chart
            var n_title = root.AddNode(pwidth,0.5);
            var title_data = new RenderItem();
            title_data.StripGridItem = null;
            title_data.Text = this.Title;
            title_data.ShapeCells = titleformat;
            n_title.Data = title_data;
            layout.PerformLayout();

            // Perform Rendering
            var dom = new VA.DOM.Document();
            foreach (var n in layout.Nodes)
            {
                if (n.Data != null)
                {
                    var s = dom.DrawRectangle(n.Rectangle);
                    
                    // Set Text
                    if (n.Data.Text !=null)
                    {
                        s.Text = this.ToUpper ? n.Data.Text.ToUpper() : n.Data.Text;
                    }

                    // Set Cells
                    if (n.Data.ShapeCells != null)
                    {
                        s.ShapeCells = n.Data.ShapeCells;
                    }

                    // draw Underline
                    if (n.Data.Underline)
                    {
                        var u = dom.DrawLine(n.Rectangle.LowerLeft, n.Rectangle.LowerRight);
                    }
                }
            }
            dom.Render(page);

        }
    }
}
