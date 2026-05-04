using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Framework
{
    [MUT.TestClass]
    public class AssemblyHooks
    {
        [MUT.AssemblyCleanup]
        public static void Cleanup()
        {
            global::VTest.Framework.VTest.TeardownVisioApplication();
        }
    }
}
