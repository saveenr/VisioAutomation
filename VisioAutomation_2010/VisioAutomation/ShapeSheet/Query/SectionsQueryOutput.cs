using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionsQueryOutput<T> : QueryOutputBase<T>
    {
        public List<SectionQueryOutput<T>> Sections { get; internal set; }

        internal SectionsQueryOutput(int shape_id, int count, List<SectionQueryOutput<T>> sections) : base(shape_id, count)
        {
            this.Sections = sections;
        }
    }
}