namespace VSamples
{
    public class SampleMethod
    {
        public string Name;
        public System.Action Method;

        public SampleMethod(string name, System.Action method)
        {
            this.Name = name;
            this.Method = method; 

        }
        public void Run()
        {
            this.Method();
        }
    }
}