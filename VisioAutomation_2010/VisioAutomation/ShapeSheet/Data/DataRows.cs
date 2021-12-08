using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Core
{
}

namespace VisioAutomation.ShapeSheet.Data
{


    public class DataRows<T> : VisioAutomation.Core.BasicList<DataRow<T>>
    {
        // Simple list of Rows


        public readonly int ShapeID;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal DataRows(int capacity) : base(capacity)
        {
            this.ShapeID = -1;
            this.SectionIndex = IVisio.VisSectionIndices.visSectionInval;
        }

        internal DataRows(int capacity, int shapeid, IVisio.VisSectionIndices section_index) : base (capacity)
        {
            this.ShapeID = shapeid;
            this.SectionIndex = section_index;
        }
    }
}