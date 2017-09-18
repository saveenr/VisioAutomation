using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.ShapeSheet;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Core.ShapeSheet
{
    [TestClass]
    public class FixedStreamBuilderTests
    {

        [TestMethod]
        public void FixedSidSrcBuilder_ThrowsException_when_not_full()
        {
            bool caught = false;

            try
            {
                var b1 = new VA.ShapeSheet.Streams.SidSrcStreamArrayBuilder(1);
                var s1 = b1.ToStreamArray();
            }
            catch (System.ArgumentException e)
            {
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not catch expected exception");
            }

            var b2 = new VA.ShapeSheet.Streams.SrcStreamArrayBuilder(1);
            b2.Add(new Src((short)0, (short)0, (short)0));
            var s2 = b2.ToStreamArray();
        }

        [TestMethod]
        public void FixedSrcBuilder_ThrowsException_when_not_full()
        {
            bool caught = false;

            try
            {
                var b1 = new VA.ShapeSheet.Streams.SrcStreamArrayBuilder(1);
                var s1 = b1.ToStreamArray();
            }
            catch (System.ArgumentException e)
            {
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not catch expected exception");
            }

            var b2 = new VA.ShapeSheet.Streams.SidSrcStreamArrayBuilder(1);
            var src = new Src((short)0, (short)0, (short)0);
            var sidsrc = new SidSrc((short)0, src);
            b2.Add(sidsrc);
            var s2 = b2.ToStreamArray();
        }

    }
}