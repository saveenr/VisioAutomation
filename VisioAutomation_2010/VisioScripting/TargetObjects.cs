using System.Collections.Generic;

namespace VisioScripting
{
    public class TargetObjects<T>
    {
        protected readonly IList<T> _items;
        public readonly bool UseContext;

        public TargetObjects()
        {
            this._items = null;
            this.UseContext = true;
        }

        public TargetObjects(IList<T> items)
        {
            this._items = items;
            this.UseContext = this._items == null;
        }

        public bool IsResolved => !this.UseContext;

    }
}