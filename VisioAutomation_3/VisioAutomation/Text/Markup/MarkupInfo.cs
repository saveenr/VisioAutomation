using System.Collections.Generic;

namespace VisioAutomation.Text.Markup
{
    public class MarkupInfo
    {
        public IList<TextRegion> FormatRegions { get; private set; }
        public IList<TextRegion> FieldRegions { get; private set; }

        public MarkupInfo()
        {
            this.FormatRegions = new List<TextRegion>();
            this.FieldRegions = new List<TextRegion>();
        }
    }
}