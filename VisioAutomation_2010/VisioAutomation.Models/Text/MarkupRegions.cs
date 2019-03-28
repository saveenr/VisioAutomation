using System.Collections.Generic;

namespace VisioAutomation.Models.Text
{
    class MarkupRegions
    {
        public List<Region> FormatRegions { get; }
        public List<Region> FieldRegions { get; }

        public MarkupRegions()
        {
            this.FormatRegions = new List<Region>();
            this.FieldRegions = new List<Region>();
        }
    }
}