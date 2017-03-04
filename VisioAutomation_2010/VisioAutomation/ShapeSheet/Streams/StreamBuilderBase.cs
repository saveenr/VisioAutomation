namespace VisioAutomation.ShapeSheet.Streams
{
    public abstract class StreamBuilderBase
    {
        public abstract short[] ToStream();

        public int Count => this._GetCount();

        protected abstract int _GetCount();

        public abstract void Clear();
    }
}



