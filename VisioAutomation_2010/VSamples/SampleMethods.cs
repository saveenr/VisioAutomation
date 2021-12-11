using System.Collections.Generic;

namespace VSamples
{
    public class SampleMethods : List<SampleMethodBase>
    {


        public SampleMethods() : base(01)
        {

        }

        public SampleMethodBase AddEx(SampleMethodBase sm)
        {
            this.Add(sm);
            return sm;
        }
    }
}