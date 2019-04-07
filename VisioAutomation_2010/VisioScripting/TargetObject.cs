namespace VisioScripting
{
    public class TargetObject<T> where  T: class
    {
        protected readonly T _item;

        public readonly bool UseContext;

        public TargetObject()
        {
            this._item = null;
            this.UseContext = true;
        }
        
        public TargetObject(T item)
        {
            this._item = item;
            this.UseContext = this._item == null;
        }
        public TargetObject(T item, bool isresolved)
        {
            this._item = item;
            this.UseContext = !isresolved;
        }

        public bool IsResolved => !this.UseContext;
    }

    public class TargetSelection 
    {

        public TargetSelection()
        {
        }

    }
}