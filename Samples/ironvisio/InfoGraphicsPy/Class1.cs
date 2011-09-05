using System;
using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using IG=InfoGraphicsPy;
using System.Linq;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace InfoGraphicsPy
{
    public class Session
    {
        public IVisio.Application app;
        
        public Session()
        {
            this.app = new IVisio.ApplicationClass();


        }

        public void NewDocument()
        {
            var docs = this.Application.Documents;
            var doc = docs.Add("");
        }

        public void NewPage()
        {
            var doc = this.Application.ActiveDocument;
            doc.Pages.Add();
        }

        public IVisio.Application Application
        {
            get { return this.app; }
        }

        public void TestDraw()
        {
            var page = this.Page;
            page.DrawRectangle(0, 0, 8.5, 11.0);

            double barwidth = 0.5;
            double hsep = 0.10;
            double vsep = 0.10;
            double maxbarheight = 4;
            double catheight = 0.5;
            var values = new double[] {1.0, 2.0, 3.0, 4.0, 5.0};

            var widths = ConstructPositions(values.Count(), barwidth, hsep);
            var heights = ConstructPositions(new[] { catheight, maxbarheight}, vsep);

            var grid = new GridLayout(widths, heights);


            int catrow = 0;
            int barrow = 2;

            var bar_rects = this.SkipOdd(grid.GetRectsInRow(barrow)).ToList();
            var cat_rects = this.SkipOdd(grid.GetRectsInRow(catrow)).ToList();

            var bar_shapes = this.DrawRects(bar_rects);
            var cat_shapes = this.DrawRects(cat_rects);


        }

        public List<IVisio.Shape> DrawRects(IEnumerable<VA.Drawing.Rectangle> rects)
        {
            var shapes = rects.Select(r => Page.DrawRectangle(r)).ToList();
            return shapes;
        }

        public static List<double> ConstructPositions(int numcols, double width, double sep)
        {
            var iwidths = new List<double>();
            for (int i = 0; i < numcols; i++)
            {
                iwidths.Add(width);
            }
            var widths = ConstructPositions(iwidths, sep);
            return widths;
        }

        public static List<double> ConstructPositions(IList<double> iwidths, double sep)
        {
            int numcols = iwidths.Count;
            var widths = new List<double>();
           
            for (int i = 0; i < numcols; i++)
            {
                widths.Add(iwidths[i]);
                if (i < numcols - 1)
                {
                    widths.Add(sep);
                }
            }
            return widths;
        }

        public IEnumerable<T> SkipOdd<T>(IEnumerable<T> items)
        {
            int i = 0;
            foreach (var item in items)
            {
                if (i % 2 == 1)
                {
                    //
                }
                else
                {
                    yield return item;
                }
                i++;
            }
            
        }


        public IVisio.Page Page
        {
            get { return this.Application.ActivePage; }
        }


        
        
    }

    public class GridLayout
    {
        public readonly List<double> Widths;
        public readonly List<double> Heights;
        public readonly List<double> Bottoms;
        public readonly List<double> Lefts;
        public readonly int ColumnCount;
        public readonly int RowCount;
            
        public GridLayout(IList<double> widths, IList<double> heights)
        {
            this.ColumnCount = widths.Count();
            this.RowCount = heights.Count();

            this.Widths = widths.ToList();
            this.Heights = heights.ToList();

            this.Bottoms = get_inc_pos(heights);
            this.Lefts = get_inc_pos(widths);

            if (this.Widths.Count() != this.Lefts.Count)
            {
                throw new Exception();
            }

            if (this.Heights.Count() != this.Bottoms.Count)
            {
                throw new Exception();
            }
        }

        public VA.Drawing.Rectangle GetRectangle(int row, int col)
        {
            var bottom = this.GetBottom(row);
            var left = this.GetLeft(col);
            var right = this.GetRight(col);
            var top = this.GetTop(row);

            return new VA.Drawing.Rectangle(left,bottom,right,top);
        }

        public IEnumerable<VA.Drawing.Rectangle> GetRectsInRow(int row)
        {
            for (int c = 0; c < this.ColumnCount; c++)
            {
                yield return this.GetRectangle(row, c);
            }
        }

        public double GetBottom(int row)
        {
            return this.Bottoms[row];
        }

        public double GetTop(int row)
        {
            return this.GetBottom(row) + this.Heights[row];
        }
        
        public double GetLeft(int col)
        {
            return this.Lefts[col];
        }

        public double GetRight(int col)
        {
            return this.GetLeft(col) + this.Widths[col];
        }

        private List<double> get_inc_pos(IList<double> lengths)
        {

            var positions = new List<double>();
            double curpos = 0.0;

            for (int i = 0; i < lengths.Count(); i++)
            {
                positions.Add(curpos);
                curpos += lengths[i];
            }
            return positions;
        }
    }
}
