using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class TextFormat_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void SetTabStops()
        {
            var no_tab_stops = new VA.Text.TabStop[] { };
            var tabstops = new[]
                               {
                                   new VA.Text.TabStop(0.5, VA.Text.TabStopAlignment.Left),
                                   new VA.Text.TabStop(1.5, VA.Text.TabStopAlignment.Right),
                                   new VA.Text.TabStop(2.5, VA.Text.TabStopAlignment.Center),
                                   new VA.Text.TabStop(4.0, VA.Text.TabStopAlignment.Decimal)
                               };

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            // shapes shoould not have any tab stops by default
            var m0 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m0.TabStops.Count);

            // clearing tab stops shoudl work even if there are no tab stops
            VA.Text.TextFormat.SetTabStops(s1, no_tab_stops);
            var m1 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m1.TabStops.Count);

            // set the 3 tab stops
            VA.Text.TextFormat.SetTabStops(s1, tabstops);

            // should have exactly the same number we set
            var m2 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(tabstops.Length, m2.TabStops.Count);
            Assert.AreEqual(0.5, tabstops[0].Position);
            Assert.AreEqual(1.5, tabstops[1].Position);
            Assert.AreEqual(2.5, tabstops[2].Position);
            Assert.AreEqual(4.0, tabstops[3].Position);
            Assert.AreEqual(VA.Text.TabStopAlignment.Left, tabstops[0].Alignment);
            Assert.AreEqual(VA.Text.TabStopAlignment.Right, tabstops[1].Alignment);
            Assert.AreEqual(VA.Text.TabStopAlignment.Center, tabstops[2].Alignment);
            Assert.AreEqual(VA.Text.TabStopAlignment.Decimal, tabstops[3].Alignment);

            // clear the tab stops
            VA.Text.TextFormat.SetTabStops(s1, no_tab_stops);
            var m3 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m3.TabStops.Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void Format1()
        {
            var text = "ABCDEFGHIJHLMNOPQRSTUVWXYZ0123456789";
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 4);
            s1.Text = text;

            var charfmt1 = new VA.Text.CharacterFormatCells();
            charfmt1.Color = "rgb(255,0,0)";
            
            VA.Text.TextFormat.SetFormat(s1, charfmt1, 0, text.Length);

            var outfmt = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(1,outfmt.CharacterFormats.Count);
            Assert.AreEqual(1, outfmt.CharacterTextRuns.Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void Format2()
        {

            var text = "ABCDEFGHIJHLMNOPQRSTUVWXYZ0123456789";
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 4);
            s1.Text = text;

            var charfmt1 = new VA.Text.CharacterFormatCells();
            charfmt1.Color = "rgb(255,0,0)";

            var charfmt2 = new VA.Text.CharacterFormatCells();
            charfmt2.Color = "rgb(0,255,0)";

            VA.Text.TextFormat.SetFormat(s1, charfmt1, 0, text.Length/2);

            var outfmt1 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(2, outfmt1.CharacterFormats.Count);
            Assert.AreEqual(2, outfmt1.CharacterTextRuns.Count);
            
            Assert.AreEqual("RGB(255,0,0)",outfmt1.CharacterFormats[0].Color.Formula);
            Assert.AreEqual("0", outfmt1.CharacterFormats[1].Color.Formula);

            VA.Text.TextFormat.SetFormat(s1, charfmt2, 5, 10);

            var outfmt2 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(4, outfmt2.CharacterFormats.Count);
            Assert.AreEqual(4, outfmt2.CharacterTextRuns.Count);

            Assert.AreEqual("RGB(255,0,0)", outfmt2.CharacterFormats[0].Color.Formula);
            Assert.AreEqual("RGB(0,255,0)", outfmt2.CharacterFormats[1].Color.Formula);
            Assert.AreEqual("RGB(255,0,0)", outfmt2.CharacterFormats[2].Color.Formula);
            Assert.AreEqual("0", outfmt2.CharacterFormats[3].Color.Formula);

            page1.Delete(0);
        }

        [TestMethod]
        public void Format3()
        {

            var text = "ABCDEFGHIJHLMNOPQRSTUVWXYZ0123456789";
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 4);
            s1.Text = text;

            var charfmt1 = new VA.Text.CharacterFormatCells();
            charfmt1.Color = "rgb(255,0,0)";

            var charfmt2 = new VA.Text.CharacterFormatCells();
            charfmt2.Color = "rgb(255,0,0)";

            VA.Text.TextFormat.SetFormat(s1, charfmt1, 0, text.Length);
            VA.Text.TextFormat.SetFormat(s1, charfmt1, 10, 20);

            var outfmt = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(3, outfmt.CharacterFormats.Count);
            Assert.AreEqual(3, outfmt.CharacterTextRuns.Count);

            Assert.AreEqual("RGB(255,0,0)", outfmt.CharacterFormats[0].Color.Formula);
            Assert.AreEqual("RGB(255,0,0)", outfmt.CharacterFormats[1].Color.Formula);
            Assert.AreEqual("RGB(255,0,0)", outfmt.CharacterFormats[2].Color.Formula);

            page1.Delete(0);
        }
    }
}