using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query 
{
    public class ShapeSectionResult<T> : Rows<T>
    {

        // for a given tuple of (shape, section) gives the rows for that tuple
        //
        // {
        //    (shapeid,sectionn)
        //    [0] = rows 0
        //    [1] = rows 1
        //    [n] = rows n
        // }

        public readonly int ShapeID;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal ShapeSectionResult(int capacity, int shapeid, IVisio.VisSectionIndices section_index) : base(capacity)
        {
            this.ShapeID = shapeid;
            this.SectionIndex = section_index;
        }
    }
}