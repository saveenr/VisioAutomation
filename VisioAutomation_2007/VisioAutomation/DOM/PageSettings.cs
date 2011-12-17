using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.DOM
{
    public class PageSettings
    {
        public Drawing.Size? Size { get; set; }
        public string Name { get; set; }
        public PageCells PageCells { get; set; }

        public PageSettings()
        {
            this.PageCells = new PageCells();
        }

        public PageSettings(VA.Drawing.Size size) :
            this()
        {
            this.Size = size;
        }

        public PageSettings(double w, double h) :
            this(new VA.Drawing.Size(w, h))
        {
        }
    }
}