using GenTreeOps;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GenTreeOps_Test
{
    public static class AssertX
    {
        public static void AssertEnter(XNode n, WalkEvent<XNode> evt)
        {
            Assert.AreEqual(n, evt.Node);
            Assert.AreEqual(WalkEvent<XNode>.WalkEventType.Enter, evt.Type);
        }

        public static void AssertExit(XNode n, WalkEvent<XNode> evt)
        {
            Assert.AreEqual(n, evt.Node);
            Assert.AreEqual(WalkEvent<XNode>.WalkEventType.Exit, evt.Type);
        }
    }
}