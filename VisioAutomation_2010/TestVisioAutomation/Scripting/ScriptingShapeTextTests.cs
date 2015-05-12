using Microsoft.VisualStudio.TestTools.UnitTesting;
	using VisioAutomation.Extensions;
using VisioAutomation.Scripting;
using VisioAutomation.Scripting.Commands;
	using IVisio = Microsoft.Office.Interop.Visio;
	using VA = VisioAutomation;
	 
	namespace TestVisioAutomation
	{
	    [TestClass]
	    public class ScriptingShapeText_Tests : VisioAutomationTest
	    {
	        [TestMethod]
	        public void Scripting_Shape_Text_Set()
	        {
	            var page1 = GetNewPage();
	            var stencil = "basic_u.vss";
	 
	            short flags = (short)IVisio.VisOpenSaveArgs.visOpenRO | (short)IVisio.VisOpenSaveArgs.visOpenDocked;
	            var app = page1.Application;
	            var documents = app.Documents;
	            var stencil_doc = documents.OpenEx(stencil, flags);
	 
	            var masters1 = stencil_doc.Masters;
	            var masters = new[] { masters1["Rounded Rectangle"], masters1["Ellipse"] };
	            var points = new[] { new VA.Drawing.Point(1, 2), new VA.Drawing.Point(3, 4) };
	            Assert.AreEqual(0, page1.Shapes.Count);
	            
	            var shapeids = page1.DropManyU(masters, points);
	            Assert.AreEqual(2, page1.Shapes.Count);
	            Assert.AreEqual(2, shapeids.Length);
	            
	            var shapes = page1.Shapes.GetShapesFromIDs(shapeids);
	            var client = this.GetScriptingClient();
	            var names = new [] { "TestName", "TestName2" };
                client.Text.Set(shapes,names);
                client.ShapeSheet.SetName(shapes,names);
	            //txtCmd.Set(shapes, names, true);
	 
	            //page1.Shapes[shapeids[0]].Text = "My Rounded Rec";
	 
	            for (int i = 0; i < page1.Shapes.Count; i++)
	            {
	                Assert.AreEqual(page1.Shapes[shapeids[i]].Name, names[i]);
	            }
	 
	            page1.Delete(0);
	        }
	    }
	}
	 