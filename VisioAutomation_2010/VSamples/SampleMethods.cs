using System.Collections.Generic;

namespace VSamples
{
    public class SampleMethods : List<SampleMethod>
    {
        public SampleMethods() : base(01)
        {
        }

        public SampleMethod AddSample(SampleMethod sm)
        {
            this.Add(sm);
            return sm;
        }
    }
}