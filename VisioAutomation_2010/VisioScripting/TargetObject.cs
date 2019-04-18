namespace VisioScripting
{
    public class TargetObject<T> where  T: class
    {
        private readonly T _item;
        public readonly bool Resolved;

        protected TargetObject()
        {
            this._item = null;
            this.Resolved = false;
        }

        protected TargetObject(T item)
        {
            this._item = item;
            this.Resolved = (item != null);
        }
        protected TargetObject(T item, bool resolved)
        {
            this._item = item;
            this.Resolved = resolved;
        }

        protected T _get_item_safe()
        {
            if (!this.Resolved)
            {
                throw new System.ArgumentException("Unresolved Target");
            }

            return this._item;
        }

    }
}