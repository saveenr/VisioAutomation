namespace VSamples
{
    public abstract class SampleMethodBase
    {
        public string GetName()
        {
            var t = this.GetType();
            return t.FullName;
        }

        public abstract void RunSample();
    }
}