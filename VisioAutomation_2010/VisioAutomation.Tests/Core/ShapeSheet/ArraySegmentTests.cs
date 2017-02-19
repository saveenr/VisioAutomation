using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Core.ShapeSheet
{
    [TestClass]
    public class ArraySegmentTests
    {

        [TestMethod]
        public void Basics()
        {
            // Check that nulls cannot be passed in
            bool caught = false;
            try
            {
                var s = new VA.Utilities.ArraySegmentBuilder<int>(null);
            }
            catch (System.ArgumentNullException e)
            {
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not catch expected exception");
            }
        }

        [TestMethod]
        public void Case1()
        {
            // Can fully accomodate an array

            var a = new int[] {1, 2, 3, 4, 5, 6, 7, 8};
            var s = new VA.Utilities.ArraySegmentBuilder<int>(a);

            var s1 = s.GetNextSegment(1);
            var s2 = s.GetNextSegment(4);
            var s3 = s.GetNextSegment(3);

            Assert.AreEqual(0, s1.Offset);
            Assert.AreEqual(1, s1.Count);

            Assert.AreEqual(1, s2.Offset);
            Assert.AreEqual(4, s2.Count);

            Assert.AreEqual(5, s3.Offset);
            Assert.AreEqual(3, s3.Count);

        }

        [TestMethod]
        public void Case2()
        {
            // Can fully accomodate an array and get multiple empty segments at end

            var a = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
            var s = new VA.Utilities.ArraySegmentBuilder<int>(a);

            var s1 = s.GetNextSegment(1);
            var s2 = s.GetNextSegment(4);
            var s3 = s.GetNextSegment(3);
            var s4 = s.GetNextSegment(0);
            var s5 = s.GetNextSegment(0);

            Assert.AreEqual(0, s1.Offset);
            Assert.AreEqual(1, s1.Count);

            Assert.AreEqual(1, s2.Offset);
            Assert.AreEqual(4, s2.Count);

            Assert.AreEqual(5, s3.Offset);
            Assert.AreEqual(3, s3.Count);

            Assert.AreEqual(5, s4.Offset);
            Assert.AreEqual(0, s4.Count);

            Assert.AreEqual(5, s5.Offset);
            Assert.AreEqual(0, s5.Count);

        }

    }
}