using System.Collections;
using System.Collections.Generic;

namespace VisioAutomation.Shapes.UserDefinedCells
{
    public class UserDefinedCellsCollection : IEnumerable<UserDefinedCellCells>
    {
        private List<UserDefinedCellCells> _list;

        internal UserDefinedCellsCollection(int capacity)
        {
            this._list = new List<UserDefinedCellCells>();
        }

        internal void Add(UserDefinedCellCells cells)
        {
            this._list.Add(cells);
        }

        public int Count => this._list.Count;

        public UserDefinedCellCells this[int index]
        {
            get { return this._list[index]; }
        }

        public IEnumerator<UserDefinedCellCells> GetEnumerator()
        {
            foreach (var i in this._list)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}