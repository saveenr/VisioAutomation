using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class OutputCollection<T> : IEnumerable<Output<T>>
    {
        private readonly List<Output<T>> _outputs;

        internal OutputCollection()
        {
            this._outputs = new List<Output<T>>();
        }

        public IEnumerator<Output<T>> GetEnumerator()
        {
            return this._outputs.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public Output<T> this[int index]
        {
            get { return this._outputs[index]; }
        }

        internal void Add(Output<T> output)
        {
            if (output == null)
            {
                throw new System.ArgumentNullException(nameof(output));
                
            }
            this._outputs.Add(output);
        }

        public int Count
        {
            get { return this._outputs.Count; }
        }
    }
}