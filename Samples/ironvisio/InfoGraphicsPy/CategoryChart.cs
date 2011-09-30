using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using VisioAutomation.DOM;
using VisioAutomation.Drawing;
using VisioAutomation.Layout.BoxHierarchy;
using BoxHierarchy=VisioAutomation.Layout.BoxHierarchy;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace InfoGraphicsPy
{

    public class CategoryItem
    {
        public string Text;
        public List<CategoryItem> Items; 
        public CategoryItem(string s)
        {
            this.Text = s;
        }
    }
    public class CategoryCell
    {
        public CategoryItem Item;
        public string XCategory;
        public string YCategory;
 
        public CategoryCell(string text, string x, string y)
        {
            this.Item = new CategoryItem(text);
            this.XCategory = x;
            this.YCategory = y;
        }
    }

    class RenderItem
    {
        public CategoryCell StripGridItem;
        public string Text ;
        public VA.DOM.ShapeCells ShapeCells;
        public bool Underline;
    }

    public class CategoryChart
    {
        public List<CategoryCell> Items;

        public string Font="Segoe UI";
        public bool ToUpper;
        public string Title = "Untitled";
        double TitleFontSize = 24;
        double CellFontSize = 8;
        double CategoryFontSize = 14;

        double CellWidth = 3.0;
        double CellVerticalSeparation = 0.125;
        double CellHeight = 0.25;
        double Indent = 2.0;
        double CategoryHeight = 0.5;
        double CellHorizontalSeparation = 0.25;

        public string CellFill = "rgb(240,240,240)";
        public string SubCellFill = "rgb(220,220,220)";

        ShapeCells titleformat = new VA.DOM.ShapeCells();
        ShapeCells cellformat = new VA.DOM.ShapeCells();
        ShapeCells subcellformat = new VA.DOM.ShapeCells();
        ShapeCells xcatformat = new VA.DOM.ShapeCells();
        ShapeCells ycatformat = new VA.DOM.ShapeCells();

        public CategoryChart()
        {
            this.Items = new List<CategoryCell>();

            titleformat.VerticalAlign = 0;
            titleformat.HAlign = 0;
            titleformat.CharSize = VA.Convert.PointsToInches(TitleFontSize);
            titleformat.LinePattern = 0;
            titleformat.LineWeight = 0;

            cellformat.VerticalAlign = 0;
            cellformat.HAlign = 0;
            cellformat.CharSize = VA.Convert.PointsToInches(CellFontSize);
            cellformat.LinePattern = 0;
            cellformat.LineWeight = 0;
            cellformat.FillForegnd = CellFill;

            subcellformat.VerticalAlign = 0;
            subcellformat.HAlign = 0;
            subcellformat.CharSize = VA.Convert.PointsToInches(CellFontSize);
            subcellformat.LinePattern = 0;
            subcellformat.LineWeight = 0;
            subcellformat.FillForegnd = SubCellFill;

            xcatformat.VerticalAlign = 2;
            xcatformat.HAlign = 1;
            xcatformat.CharSize = VA.Convert.PointsToInches(CategoryFontSize);
            xcatformat.LinePattern = 0;
            xcatformat.LineWeight = 0;
            xcatformat.CharStyle = ((int)VA.Text.CharStyle.Bold);

            ycatformat.VerticalAlign = 2;
            ycatformat.HAlign = 0;
            ycatformat.CharSize = VA.Convert.PointsToInches(CategoryFontSize);
            ycatformat.LinePattern = 0;
            ycatformat.LineWeight = 0;
            ycatformat.CharStyle = ((int)VA.Text.CharStyle.Bold);

        }

        public CategoryCell Add(string text, string xcat, string ycat)
        {
            var item = new CategoryCell(text,xcat,ycat);
            this.Items.Add(item);
            return item;
        }

        public CategoryCell Add(string text, string xcat, string ycat, IList<string> subitems)
        {
            var item = new CategoryCell(text, xcat, ycat);

            item.Item.Items = subitems.Select(t=>new CategoryItem(t)).ToList();
            this.Items.Add(item);
            return item;
        }
        public void Render(IVisio.Page page)
        {
            var xcats = this.Items.Select(i => i.XCategory).Distinct().ToList();
            var ycats = this.Items.Select(i => i.YCategory).Distinct().ToList();

            int cols = xcats.Count();
            int rows = ycats.Count();
            double pwidth = cols * CellWidth + (System.Math.Max(0, cols - 1) * 0.25) + Indent + 0.25;

            var layout = new BoxHierarchy.BoxHierarchyLayout<RenderItem>();
            layout.LayoutOptions.Origin = new VA.Drawing.Point(0,10);
            layout.LayoutOptions.DefaultHeight = 0.25;
            var root = layout.Root;
            root.Direction = BoxHierarchy.LayoutDirection.Vertical;
            root.ChildSeparation = 0.125;

            foreach (int row in Enumerable.Range(0, rows))
            {

                var n_ycat_row = root.AddNode(BoxHierarchy.LayoutDirection.Horizonal);
                n_ycat_row.ChildSeparation = CellHorizontalSeparation;

                // -- add indent
                n_ycat_row.AddNode(Indent, 0.25);

                foreach (int col in Enumerable.Range(0, cols))
                {
                    var n_row_col = n_ycat_row.AddNode(CellWidth, 0.25);

                    // ---
                    n_row_col.Direction = BoxHierarchy.LayoutDirection.Vertical;
                    n_row_col.AlignmentVertical = AlignmentVertical.Top;
                    n_row_col.ChildSeparation = CellVerticalSeparation;
                    var items_for_cells = this.Items.Where(i => i.XCategory == xcats[col] && i.YCategory == ycats[row]);
                    foreach (var cell_item in items_for_cells)
                    {
                        draw_cell(cell_item, n_row_col);
                    }
                }

                var n_ycat_title = root.AddNode(pwidth, CategoryHeight);
                var ycat_data = new RenderItem();
                ycat_data.StripGridItem = null;
                ycat_data.Text = ycats[row];
                ycat_data.ShapeCells = ycatformat;
                ycat_data.Underline = true;
                n_ycat_title.Data = ycat_data;
            }

            var n_xcatlabels = root.AddNode(BoxHierarchy.LayoutDirection.Horizonal);
            n_xcatlabels.ChildSeparation = CellHorizontalSeparation;

            // Add indent
            n_xcatlabels.AddNode(Indent, 0.25);

            // Add XCategory labels
            foreach (int col in Enumerable.Range(0, cols))
            {
                var n = n_xcatlabels.AddNode(CellWidth, 0.5);
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

            Render(page, layout);
        }

        private void draw_cell(CategoryCell cell_item, Node<RenderItem> n_row_col)
        {
            var n_cell = n_row_col.AddNode(CellWidth, CellHeight);
            n_cell.ChildSeparation = CellVerticalSeparation/2;
            
            var cell_data = new RenderItem();
            cell_data.StripGridItem = cell_item;
            cell_data.Text = cell_item.Item.Text;
            cell_data.ShapeCells = cellformat;
            n_cell.Data = cell_data;
            
            if (cell_item.Item.Items != null)
            {
                foreach (var sub_cat_items in cell_item.Item.Items)
                {
                    var subn_cell = n_cell.AddNode(CellWidth, CellHeight);
                    var subcell_data = new RenderItem();
                    subcell_data.StripGridItem = null;
                    subcell_data.Text = sub_cat_items.Text;
                    subcell_data.ShapeCells = subcellformat;
                    subn_cell.Data = subcell_data;
                }
                n_cell.AddNode(null, 0.25);
            }
        }

        private void Render(Page page, BoxHierarchyLayout<RenderItem> layout)
        {
            layout.PerformLayout();
            var doc = page.Document;
            var fonts = doc.Fonts;
            var default_font = fonts[this.Font];
            int default_font_id = default_font.ID;
            // Perform Rendering
            var dom = new VA.DOM.Document();
            foreach (var n in layout.Nodes)
            {
                if (n.Data != null)
                {
                    var s = dom.DrawRectangle(n.ReservedRectangle);

                    // Set Text
                    if (n.Data.Text != null)
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

                    n.Data.ShapeCells.CharFont = default_font_id;
                }
            }
            dom.Render(page);
        }

        public static CategoryChart FromCSV(string title, string text)
        {
            var chart = new CategoryChart();
            chart.Title = title;
            foreach (var line in text.Split(new char[] { '\n' }))
            {
                var sline = line.Trim();
                if (sline.Length < 1)
                {
                    continue;
                }

                var tokens = line.Split(new char[] { ',' });
                if (tokens.Length < 3)
                {
                    throw new System.Exception("Not enough tokens in line");
                }

                string xcat = tokens[0];
                string ycat = tokens[1];
                string item = tokens[2];
                string[] subitems = tokens.Length >= 4
                                        ? tokens[3].Split(new char[] { '|' }).Select(s => s.Trim()).Where(s => s.Length > 0).
                                              ToArray()
                                        : null;
                if (subitems == null)
                {
                    chart.Add(item, xcat, ycat);
                }
                else
                {
                    chart.Add(item, xcat, ycat, subitems);
                }
            }

            return chart;
        }

    }
}
