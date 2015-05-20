using System;
using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;

namespace InfoGraphicsPy
{
    public enum StartOfWeek
    {
        Sunday,
        Monday
    }

    public class MonthGrid
    {
        public MonthYear MonthYear { get; private set; }

        public StartOfWeek StartOfWeek;

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

        private Dictionary<DayOfWeek, int> GetDayOfWeekDic()
        {
            if (this.StartOfWeek == StartOfWeek.Sunday)
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
            else if (this.StartOfWeek == StartOfWeek.Sunday)
            {
                var dic = new Dictionary<System.DayOfWeek, int>
                          {
                              {System.DayOfWeek.Sunday, 6},
                              {System.DayOfWeek.Monday, 0},
                              {System.DayOfWeek.Tuesday, 1},
                              {System.DayOfWeek.Wednesday, 2},
                              {System.DayOfWeek.Thursday, 3},
                              {System.DayOfWeek.Friday, 4},
                              {System.DayOfWeek.Saturday, 5}
                          };
                return dic;
            }
            else
            {
                throw new Exception();
            }
        }
    }
}