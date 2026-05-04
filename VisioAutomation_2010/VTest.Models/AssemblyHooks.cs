using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Models
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
