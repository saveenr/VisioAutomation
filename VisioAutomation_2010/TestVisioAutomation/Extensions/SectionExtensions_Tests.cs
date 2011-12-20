using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;
using System.Linq;
using System.Collections.Generic;

namespace TestVisioAutomation
{
    [TestClass]
    public class SectionExtensions_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void CreatePage()
        {
            var page1 = GetNewPage();
            var doc1 = page1.Document;

            // Create a simple shape with text that has multiple character formatting rows
            // Then format the text and all character rows should be altered

            var shape0 = page1.DrawRectangle(1, 1, 3, 3);

            shape0.Text = TestCommon.Helper.LoremIpsumText;

            var fmt0 = new VA.Text.CharacterFormatCells();
            var pts_10 = VA.Convert.PointsToInches(10);
            fmt0.Size = pts_10;

            var fmt1 = new VA.Text.CharacterFormatCells();
            var pts_6 = VA.Convert.PointsToInches(6);
            fmt1.Size = pts_6;

            var fmt2 = new VA.Text.CharacterFormatCells();
            var pts_18 = VA.Convert.PointsToInches(18);
            fmt2.Size = pts_18;

            var fmt3 = new VA.Text.CharacterFormatCells();
            var pts_9 = VA.Convert.PointsToInches(9);
            fmt3.Size = pts_9;

            VisioAutomation.Text.TextFormat.SetFormat(shape0, fmt0);
            VisioAutomation.Text.TextFormat.SetFormat(shape0, fmt1, 10, 20);
            VisioAutomation.Text.TextFormat.SetFormat(shape0, fmt2, 30, 40);


            var section = shape0.Section[(short) IVisio.VisSectionIndices.visSectionCharacter];

            var expected = new List<IVisio.Row>();
            for (int i = 0; i < section.Count; i++)
            {
                expected.Add( section[(short)i]);
            }
            var actual = section.AsEnumerable().ToList();

            Assert.AreEqual(expected.Count, actual.Count);

            Assert.AreEqual(section[0].Index, expected[0].Index);
            Assert.AreEqual(section[1].Index, expected[1].Index);
            Assert.AreEqual(section[2].Index, expected[2].Index);
            Assert.AreEqual(section[3].Index, expected[3].Index);

            doc1.Close(true);
        }

    }
}