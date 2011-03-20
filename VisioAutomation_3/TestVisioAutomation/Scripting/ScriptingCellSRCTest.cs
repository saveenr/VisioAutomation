using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass]
    public class ScriptingSRCTest : VisioAutomationTest
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void SpotCheck1()
        {
            var c1 = VA.ShapeSheet.ShapeSheetHelper.TryGetSRCFromName("EndArrow").Value;
            var c2 = VA.ShapeSheet.SRCConstants.EndArrow;

            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(c2, c1);
        }
    }
}