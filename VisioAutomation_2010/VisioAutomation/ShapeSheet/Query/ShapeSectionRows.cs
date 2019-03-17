using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query 
{
    public class ShapeSectionRows<T> : IEnumerable<ShapeCellsRow<T>>
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
        private readonly RowList<T> Rows;

        internal ShapeSectionRows(int capacity, int shapeid, IVisio.VisSectionIndices section_index)
        {
            this.ShapeID = shapeid;
            this.Rows = new RowList<T>(capacity);
            this.SectionIndex = section_index;
        }


    public IEnumerator<ShapeCellsRow<T>> GetEnumerator()
    {
        return this.Rows.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    internal void Add(ShapeCellsRow<T> r)
    {
        this.Rows.Add(r);
    }

    internal void AddRange(IEnumerable<ShapeCellsRow<T>> rows)
    {
        this.Rows.AddRange(rows);
    }

    public int Count => this.Rows.Count;

    public ShapeCellsRow<T> this[int index] => this.Rows[index];

}
}