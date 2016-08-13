using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Outputs
{
    public class ListOutput<T> : IEnumerable<Output<T>>
    {
        private readonly List<Output<T>> _outputs;

        internal ListOutput()
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

        public Output<T> this[int index] => this._outputs[index];

        internal void Add(Output<T> item)
        {
            this._outputs.Add(item);
        }

        public int Count => this._outputs.Count;
    }
}