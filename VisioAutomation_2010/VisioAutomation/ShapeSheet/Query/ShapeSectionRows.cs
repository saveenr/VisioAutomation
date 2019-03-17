using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query 
{
    public class ShapeSectionRows<T> : IEnumerable<Row<T>>
    {
        public readonly int ShapeID;
        public readonly IVisio.VisSectionIndices SectionIndex;
        private readonly RowList<T> Rows;

        internal ShapeSectionRows(int shapeid, int capacity, IVisio.VisSectionIndices section_index)
        {
            this.ShapeID = shapeid;
            this.Rows = new RowList<T>(capacity);
            this.SectionIndex = section_index;
        }


    public IEnumerator<Row<T>> GetEnumerator()
    {
        return this.Rows.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    internal void Add(Row<T> r)
    {
        this.Rows.Add(r);
    }

    internal void AddRange(IEnumerable<Row<T>> rows)
    {
        this.Rows.AddRange(rows);
    }

    public int Count => this.Rows.Count;

    public Row<T> this[int index] => this.Rows[index];

}
}