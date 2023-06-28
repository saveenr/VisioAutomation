namespace VSamples
{
    public abstract class SampleMethod
    {
        private string name;

        public string Name
        {
            get
            {
                if (name == null)
                {
                    this.name = this.GetType().FullName;
                }

                return this.name;
            }
        }

        public abstract void Run();
    }
}