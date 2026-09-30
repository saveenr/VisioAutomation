using System;
using System.Linq;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.PowerShell
{
    [MUT.TestClass]
    public class SessionTests
    {
        [MUT.TestMethod]
        public void Runspace_UsesTheTestBuildWithoutModuleAutoloading()
        {
            using (var session = new VisioPSSession())
            {
                var paths = session.InvokeScriptStrict<string>(
                    "$PSModuleAutoLoadingPreference = 'None'; " +
                    "(Get-Command New-VisioDocument -ErrorAction Stop).ImplementingType.Assembly.Location");

                MUT.Assert.AreEqual(1, paths.Count);
                MUT.Assert.IsTrue(string.Equals(
                    typeof(VisioPowerShell.Commands.VisioCmdlet).Assembly.Location,
                    paths.Single(), StringComparison.OrdinalIgnoreCase),
                    "The runspace must use the same VisioPS assembly as directly invoked cmdlets.");
            }
        }
    }
}
