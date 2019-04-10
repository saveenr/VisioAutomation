using System.Collections.Generic;

namespace VisioScripting
{
    public class TargetObjects<T>
    {
        protected readonly IList<T> _items;
        public readonly bool Resolved;

        public TargetObjects()
        {
            this._items = null;
            this.Resolved = true;
        }

        public TargetObjects(IList<T> items)
        {
            this._items = items;
            this.Resolved = (items == null);
        }

    }
}