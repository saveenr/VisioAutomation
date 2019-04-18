using System.Collections.Generic;

namespace VisioScripting
{
    public class TargetObjects<T>
    {
        private readonly IList<T> _items;
        public readonly bool Resolved;

        protected TargetObjects()
        {
            this._items = null;
            this.Resolved = false;
        }

        protected TargetObjects(IList<T> items)
        {
            this._items = items;
            this.Resolved = (items != null);
        }

        protected IList<T> _get_items_safe()
        {
            if (!this.Resolved)
            {
                throw new System.ArgumentException("Unresolved Target Collection");
            }
            return this._items;
        }

        internal IList<T> _get_items_unsafe()
        {
            return this._items;
        }
    }
}