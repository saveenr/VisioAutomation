using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    public class MultiSectionOutput<T> : OutputBase
    {
        public SectionOutputList<T> Sections { get; internal set; }

        internal MultiSectionOutput(int shape_id, int count, SectionOutputList<T> sections) : base(shape_id, count)
        {
            this.Sections = sections;
        }
    }
}