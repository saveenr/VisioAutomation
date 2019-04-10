namespace VisioScripting
{
    public class TargetObject<T> where  T: class
    {
        protected readonly T _item;
        public readonly bool Resolved;

        public TargetObject()
        {
            this._item = null;
            this.Resolved = false;
        }
        
        public TargetObject(T item)
        {
            this._item = item;
            this.Resolved = (item !=null);
        }
        public TargetObject(T item, bool resolved)
        {
            this._item = item;
            this.Resolved = resolved;
        }

    }
}