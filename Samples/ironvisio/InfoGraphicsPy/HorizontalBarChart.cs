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
    public class HorizontalBarChart : GridChart
    {
        public string[] CategoryLabels;
        public DataPoints DataPoints;
        new double CellHeight = 0.25;

        public HorizontalBarChart(DataPoints dps, string[] cats)
        {
            this.DataPoints = dps;
            this.CategoryLabels = cats;
            
        }

        public void Draw(Session session)
        {
            var normalized_values = this.DataPoints.GetNormalizedValues();
            var heights = DOMUTil.ConstructPositions(this.DataPoints.Count(), CellHeight, this.VerticalSeparation);
            var widths = DOMUTil.ConstructPositions(new[] { this.CategoryLabelHeight, this.CellWidth }, this.HorizontalSeparation);
            var grid = new GridLayout(widths, heights);

            int catcol = 0;
            int barcol = 2;

            var content_rects = this.SkipOdd(grid.GetRectsInCol(barcol)).ToList();

            var dom = new VA.DOM.Document();
            dom.ResolveVisioShapeObjects = true;

            var bar_rects = new List<VA.Drawing.Rectangle>(content_rects.Count);
            for (int i = 0; i < content_rects.Count; i++)
            {
                var r = content_rects[i];
                dom.DrawRectangle(r);
                var size = new VA.Drawing.Size(normalized_values[i] * r.Width, this.CellHeight);
                var bar_rect = new VA.Drawing.Rectangle(r.LowerLeft, size);
                bar_rects.Add(bar_rect);
            }

            var cat_rects = this.SkipOdd(grid.GetRectsInCol(catcol)).ToList();

            var bar_shapes = DOMUTil.DrawRects(dom, bar_rects, session.MasterRectangle);
            var cat_shapes = DOMUTil.DrawRects(dom, cat_rects, session.MasterRectangle);

            for (int i = 0; i < this.DataPoints.Count; i++)
            {
                bar_shapes[i].Text = this.DataPoints[i].Text.ToString();
                cat_shapes[i].Text = this.CategoryLabels[i];
            }

            foreach (var shape in bar_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = this.ValueFillColor;
                cells.LineColor = this.LineLightBorder;

            }

            foreach (var shape in cat_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillPattern = this.CategoryFillPattern;
                cells.LineWeight = this.CategoryLineWeight;
                cells.LinePattern = this.CategoryLinePattern;
            }
            dom.Render(session.Page);
        }
    }

    public class MonthYear
    {
        public int Month { get; private set; }
        public int Year { get; private set; }

        public MonthYear(int year, int month)
        {
            this.Month = month;
            this.Year = year;
        }

        public MonthYear(System.DateTime datetime)
        {
            this.Month = datetime.Month;
            this.Year = datetime.Year;
        }

        public System.DateTime FirstDay
        {
            get { return new System.DateTime(this.Year, this.Month, 1); }
        }

        public int DaysInMonth
        {
            get { return DateTime.DaysInMonth(this.Year, this.Month); }
        }

        public System.DateTime LastDay
        {
            get
            {
                return new System.DateTime(this.Year, this.Month, this.DaysInMonth);
            }
        }
    }

    public class MonthGrid
    {
        public MonthYear MonthYear { get; private set; }

        public MonthGrid(int year, int month)
        {
            this.MonthYear = new MonthYear(year,month);
        }

        public void Render(IVisio.Page page)
        {

            var dic = GetDayOfWeekDic();

            // calcualte actual days in month
            int weekday = 0 + dic[this.MonthYear.FirstDay.DayOfWeek];
            int week = 0;
          
            foreach (int day in Enumerable.Range(0,this.MonthYear.DaysInMonth))
            {
                double x = 0.0 + weekday*1.0;
                double y = 6.0 - week*1.0;
                var shape = page.DrawRectangle(x, y, x + 0.9, y + 0.9);

                weekday++;
                if (weekday >= 7)
                {
                    week++;
                    weekday = 0;
                }

                shape.Text = string.Format("{0}", new System.DateTime(this.MonthYear.Year, this.MonthYear.Month, day+1));
            }
        }

        private static Dictionary<DayOfWeek, int> GetDayOfWeekDic()
        {
            var dic = new Dictionary<System.DayOfWeek, int>
                          {
                              {System.DayOfWeek.Sunday, 0},
                              {System.DayOfWeek.Monday, 1},
                              {System.DayOfWeek.Tuesday, 2},
                              {System.DayOfWeek.Wednesday, 3},
                              {System.DayOfWeek.Thursday, 4},
                              {System.DayOfWeek.Friday, 5},
                              {System.DayOfWeek.Saturday, 6}
                          };
            return dic;
        }
    }
}
