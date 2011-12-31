using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class CellGroups : VisioAutomationTest
    {
        [TestMethod]
        public void VerifyCellGroupMembers()
        {
            var cells = new VA.Format.ShapeFormatCells();
            var members = cells.GetCellMembers().ToDictionary(i=>i.Name);

            var fillforegnd = members["FillForegnd"];
            Assert.AreEqual(typeof(int),fillforegnd.DataType);
            Assert.AreEqual(25,members.Count);
        }

    }
}
