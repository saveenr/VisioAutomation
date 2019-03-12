using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Utilities;
using VA = VisioAutomation;


namespace VisioAutomation_Tests.Utilities
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
                var s = new VA.ShapeSheet.Internal.ArraySegmentReader<int>(null);
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
        public void Segment_read_full_array_exactly()
        {
            // Can fully accomodate an array

            var a = new int[] {1, 2, 3, 4, 5, 6, 7, 8};
            var s = new VA.ShapeSheet.Internal.ArraySegmentReader<int>(a);

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
        public void Segment_read_full_array_exactly_multiple_empty_at_end()
        {
            // Can fully accomodate an array and get multiple empty segments at end

            var a = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
            var s = new VA.ShapeSheet.Internal.ArraySegmentReader<int>(a);

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

            Assert.AreEqual(8, s4.Offset);
            Assert.AreEqual(0, s4.Count);

            Assert.AreEqual(8, s5.Offset);
            Assert.AreEqual(0, s5.Count);

        }

        [TestMethod]
        public void Segment_error_if_asked_to_produce_too_much_1()
        {
            // fails if asks too much - current position is in middle of array

            var a = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
            var s = new VA.ShapeSheet.Internal.ArraySegmentReader<int>(a);
            var s1 = s.GetNextSegment(4);

            Assert.AreEqual(0, s1.Offset);
            Assert.AreEqual(4, s1.Count);

            CheckOverflow(s, 5);
        }

        [TestMethod]
        public void Segment_error_if_asked_to_produce_too_much_2()
        {
            // fails if asks too much - current position is at start middle strt of array after asking for empty segment

            var a = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
            var s = new VA.ShapeSheet.Internal.ArraySegmentReader<int>(a);
            var s1 = s.GetNextSegment(0);

            Assert.AreEqual(0, s1.Offset);
            Assert.AreEqual(0, s1.Count);

            CheckOverflow(s, 9);
        }

        [TestMethod]
        public void Segment_ask_for_entire_array_at_once()
        {
            // fails if asks too much - current position is at start middle of array

            var a = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
            var s = new VA.ShapeSheet.Internal.ArraySegmentReader<int>(a);
            var s1 = s.GetNextSegment(8);

            Assert.AreEqual(0, s1.Offset);
            Assert.AreEqual(8, s1.Count);

            CheckOverflow(s, 1);
        }


        private static void CheckOverflow(VA.ShapeSheet.Internal.ArraySegmentReader<int> s, int size)
        {
            bool caught = false;
            try
            {
                var s2 = s.GetNextSegment(size);
            }
            catch (System.ArgumentOutOfRangeException e)
            {
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not catch expected exception");
            }
        }
    }
}