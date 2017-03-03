using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.ShapeSheet;
using VisioAutomation.Utilities;
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
                var b1 = new VA.ShapeSheet.Streams.FixedSidSrcStreamBuilder(1);
                var s1 = b1.ToStream();
            }
            catch (System.ArgumentException e)
            {
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not catch expected exception");
            }

            var b2 = new VA.ShapeSheet.Streams.FixedSrcStreamBuilder(1);
            b2.Add(new Src((short)0, (short)0, (short)0));
            var s2 = b2.ToStream();
        }

        [TestMethod]
        public void FixedSrcBuilder_ThrowsException_when_not_full()
        {
            bool caught = false;

            try
            {
                var b1 = new VA.ShapeSheet.Streams.FixedSrcStreamBuilder(1);
                var s1 = b1.ToStream();
            }
            catch (System.ArgumentException e)
            {
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not catch expected exception");
            }

            var b2 = new VA.ShapeSheet.Streams.FixedSidSrcStreamBuilder(1);
            b2.Add(new SidSrc((short)0, (short)0, (short)0, (short)0));
            var s2 = b2.ToStream();
        }

    }
}