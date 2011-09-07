using System;
using System.Collections;
using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using IG=InfoGraphicsPy;
using System.Linq;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace InfoGraphicsPy
{
   
    public class GridLayout
    {
        public readonly List<double> Widths;
        public readonly List<double> Heights;
        public readonly List<double> Bottoms;
        public readonly List<double> Lefts;
        public readonly int ColumnCount;
        public readonly int RowCount;

        public enum VerticalDirection
        {
            BottomToTop,
            TopToBottom
        }
            
        public GridLayout(IList<double> widths, IList<double> heights)
        {
            this.ColumnCount = widths.Count();
            this.RowCount = heights.Count();

            this.Widths = widths.ToList();
            this.Heights = heights.ToList();

            this.Bottoms = get_inc_pos(0.0,heights);
            this.Lefts = get_inc_pos(0.0, widths);

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

        private List<double> get_inc_pos(double startpos, IList<double> lengths)
        {
            var positions = new List<double>(lengths.Count);
            double curpos = startpos;

            for (int i = 0; i < lengths.Count(); i++)
            {
                positions.Add(curpos);
                curpos += lengths[i];
            }
            return positions;
        }
    }
}
