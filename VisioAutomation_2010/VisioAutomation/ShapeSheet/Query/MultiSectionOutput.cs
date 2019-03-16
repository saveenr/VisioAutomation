using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class MultiSectionOutput<T> : OutputBase<T>
    {
        public List<SectionOutput<T>> Sections { get; internal set; }

        internal MultiSectionOutput(int shape_id, int count, List<SectionOutput<T>> sections) : base(shape_id, count)
        {
            this.Sections = sections;
        }
    }
}