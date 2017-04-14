using System.Collections.Generic;

namespace VisioAutomation.Models.Text
{
    class MarkupRegions
    {
        public List<Region> FormatRegions { get; private set; }
        public List<Region> FieldRegions { get; private set; }

        public MarkupRegions()
        {
            this.FormatRegions = new List<Region>();
            this.FieldRegions = new List<Region>();
        }
    }
}