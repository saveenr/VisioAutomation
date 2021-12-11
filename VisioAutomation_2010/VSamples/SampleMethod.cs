using System.Collections.Generic;
using System.ComponentModel;

namespace VSamples
{

    public class SampleMethods : List<SampleMethod>
    {


        public SampleMethods() : base(01)
        {

        }

        public SampleMethod Add(string name, System.Action method)
        {
            var m = new SampleMethod(name, method);
            this.Add(m);
            return m;
        }
    }

    public class SampleMethodBase
    {
        public string GetName()
        {
            var t = this.GetType();
            return t.FullName;
        }
    }
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