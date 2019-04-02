using System.Collections.Generic;

namespace VisioScripting.Models
{
    public class TargetObjects<T>
    {
        public readonly IList<T> Items;
        public readonly bool UseContext;

        public TargetObjects()
        {
            this.Items = null;
            this.UseContext = true;
        }

        public TargetObjects(IList<T> items)
        {
            this.Items = items;
            this.UseContext = this.Items == null;
        }

        public bool IsResolved => !this.UseContext;

    }
}