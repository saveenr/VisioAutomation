namespace VisioScripting
{
    public class TargetObject<T> where  T: class
    {
        public readonly T Item;
        public readonly bool UseContext;

        public TargetObject()
        {
            this.Item = null;
            this.UseContext = true;
        }
        
        public TargetObject(T item)
        {
            this.Item = item;
            this.UseContext = this.Item == null;
        }
        public TargetObject(T item, bool isresolved)
        {
            this.Item = item;
            this.UseContext = !isresolved;
        }

        public bool IsResolved => !this.UseContext;
    }
}