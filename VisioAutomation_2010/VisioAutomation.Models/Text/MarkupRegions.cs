using System.Collections.Generic;

namespace VisioAutomation.Models.Text
{
    class MarkupRegions
    {
        public List<TextRegion> FormatRegions { get; private set; }
        public List<TextRegion> FieldRegions { get; private set; }

        public MarkupRegions()
        {
            this.FormatRegions = new List<TextRegion>();
            this.FieldRegions = new List<TextRegion>();
        }
    }
}