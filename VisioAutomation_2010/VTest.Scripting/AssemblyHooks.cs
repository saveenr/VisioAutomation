using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class AssemblyHooks
    {
        [MUT.AssemblyCleanup]
        public static void Cleanup()
        {
            VTest.Framework.VTest.TeardownVisioApplication();
        }
    }
}
