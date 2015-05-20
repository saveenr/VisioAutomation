using System;

namespace InfoGraphicsPy
{
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
}