using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class MultiSectionQueryOutput<T> : QueryOutputBase<T>
    {
        public List<SectionQueryOutput<T>> Sections { get; internal set; }

        internal MultiSectionQueryOutput(int shape_id, int count, List<SectionQueryOutput<T>> sections) : base(shape_id, count)
        {
            this.Sections = sections;
        }
    }
}