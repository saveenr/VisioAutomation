using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;


namespace VTest.Utilities
{
    [MUT.TestClass]
    public class ArraySegmentTests
    {

        [MUT.TestMethod]
        public void Basics()
        {
            // Check that nulls cannot be passed in
            bool caught = false;
            try
            {
                var s = new VA.Internal.ArraySegmentEnumerator<int>(null);
            }
            catch (System.ArgumentNullException)
            {
                caught = true;
            }

            if (!caught)
            {
                MUT.Assert.Fail("Did not catch expected exception");
            }
        }

        [MUT.TestMethod]
        public void Segment_read_full_array_exactly()
        {
            // Can fully accomodate an array

            var a = new int[] {1, 2, 3, 4, 5, 6, 7, 8};
            var s = new VA.Internal.ArraySegmentEnumerator<int>(a);

            var s1 = s.GetNextSegment(1);
            var s2 = s.GetNextSegment(4);
            var s3 = s.GetNextSegment(3);

            MUT.Assert.AreEqual(0, s1.Offset);
            MUT.Assert.AreEqual(1, s1.Count);

            MUT.Assert.AreEqual(1, s2.Offset);
            MUT.Assert.AreEqual(4, s2.Count);

            MUT.Assert.AreEqual(5, s3.Offset);
            MUT.Assert.AreEqual(3, s3.Count);

        }

        [MUT.TestMethod]
        public void Segment_read_full_array_exactly_multiple_empty_at_end()
        {
            // Can fully accomodate an array and get multiple empty segments at end

            var a = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
            var s = new VA.Internal.ArraySegmentEnumerator<int>(a);

            var s1 = s.GetNextSegment(1);
            var s2 = s.GetNextSegment(4);
            var s3 = s.GetNextSegment(3);
            var s4 = s.GetNextSegment(0);
            var s5 = s.GetNextSegment(0);

            MUT.Assert.AreEqual(0, s1.Offset);
            MUT.Assert.AreEqual(1, s1.Count);

            MUT.Assert.AreEqual(1, s2.Offset);
            MUT.Assert.AreEqual(4, s2.Count);

            MUT.Assert.AreEqual(5, s3.Offset);
            MUT.Assert.AreEqual(3, s3.Count);

            MUT.Assert.AreEqual(8, s4.Offset);
            MUT.Assert.AreEqual(0, s4.Count);

            MUT.Assert.AreEqual(8, s5.Offset);
            MUT.Assert.AreEqual(0, s5.Count);

        }

        [MUT.TestMethod]
        public void Segment_error_if_asked_to_produce_too_much_1()
        {
            // fails if asks too much - current position is in middle of array

            var a = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
            var s = new VA.Internal.ArraySegmentEnumerator<int>(a);
            var s1 = s.GetNextSegment(4);

            MUT.Assert.AreEqual(0, s1.Offset);
            MUT.Assert.AreEqual(4, s1.Count);

            _check_overflow(s, 5);
        }

        [MUT.TestMethod]
        public void Segment_error_if_asked_to_produce_too_much_2()
        {
            // fails if asks too much - current position is at start middle strt of array after asking for empty segment

            var a = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
            var s = new VA.Internal.ArraySegmentEnumerator<int>(a);
            var s1 = s.GetNextSegment(0);

            MUT.Assert.AreEqual(0, s1.Offset);
            MUT.Assert.AreEqual(0, s1.Count);

            _check_overflow(s, 9);
        }

        [MUT.TestMethod]
        public void Segment_ask_for_entire_array_at_once()
        {
            // fails if asks too much - current position is at start middle of array

            var a = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
            var s = new VA.Internal.ArraySegmentEnumerator<int>(a);
            var s1 = s.GetNextSegment(8);

            MUT.Assert.AreEqual(0, s1.Offset);
            MUT.Assert.AreEqual(8, s1.Count);

            _check_overflow(s, 1);
        }


        private static void _check_overflow(VA.Internal.ArraySegmentEnumerator<int> s, int size)
        {
            bool caught = false;
            try
            {
                var s2 = s.GetNextSegment(size);
            }
            catch (System.ArgumentOutOfRangeException)
            {
                caught = true;
            }

            if (!caught)
            {
                MUT.Assert.Fail("Did not catch expected exception");
            }
        }
    }
}