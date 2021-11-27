using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VA=VisioAutomation;

namespace VTest.Core.Extensions
{
    [MUT.TestClass]
    public class ApplicationTests : VisioAutomationTest
    {
        /*
        
        USAGE: When using the UndoScope pattern, the Undo()
        must occur after the } corresponding to the undo scope, not inside of it
                
        CORRECT USAGE EXAMPLE:
        
            using (var undoscope1 = app.CreateUndoScope())
            {
                // do something
            }
            app.Undo();

        INCORRECT CORRECT USAGE EXAMPLE:
        
            using (var undoscope1 = app.CreateUndoScope())
            {
                // do something
                app.Undo();
            }

        */

        [MUT.TestMethod]
        public void Application_UndoScope_Simple()
        {
            var page1 = this.GetNewPage();
            var app = page1.Application;

            MUT.Assert.AreEqual(0, page1.Shapes.Count);

            // create a shape without undoing it
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 2, 2);
                MUT.Assert.AreEqual(1, page1.Shapes.Count);
            }
            MUT.Assert.AreEqual(1, page1.Shapes.Count);

            // create a shape and undo it
            using (var undoscope1 = new VA.Application.UndoScope(app, "UndoScope2"))
            {
                var s1 = page1.DrawRectangle(1, 1, 3, 3);
                MUT.Assert.AreEqual(2, page1.Shapes.Count);
            }
            app.Undo();

            MUT.Assert.AreEqual(1, page1.Shapes.Count);
            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Application_UndoScope_NestedInner()
        {
            var page1 = this.GetNewPage();
            var app = page1.Application;

            MUT.Assert.AreEqual(0, page1.Shapes.Count);

            // Test that inner undo doesn't affect outer
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 2, 2);
                MUT.Assert.AreEqual(1, page1.Shapes.Count);

                // create a shape and undo it
                using (var undoscope1 = new VA.Application.UndoScope(app, "UndoScope2"))
                {
                    var s2 = page1.DrawRectangle(1, 1, 3, 3);
                    MUT.Assert.AreEqual(2, page1.Shapes.Count);
                }
                app.Undo();
            }
            MUT.Assert.AreEqual(1, page1.Shapes.Count);
            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Application_UndoScope_NestedOuter()
        {
            var page1 = this.GetNewPage();
            var app = page1.Application;

            MUT.Assert.AreEqual(0, page1.Shapes.Count);

            // Test that outter does affect inner
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 2, 2);
                MUT.Assert.AreEqual(1, page1.Shapes.Count);

                // create a shape and undo it
                using (var undoscope1 = new VA.Application.UndoScope(app, "UndoScope2"))
                {
                    var s2 = page1.DrawRectangle(1, 1, 3, 3);
                    MUT.Assert.AreEqual(2, page1.Shapes.Count);
                }
            }
            app.Undo();

            MUT.Assert.AreEqual(0, page1.Shapes.Count);
            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Application_UndoScope_Abort()
        {
            var page1 = this.GetNewPage();
            var app = page1.Application;

            MUT.Assert.AreEqual(0, page1.Shapes.Count);

            // create a shape without undoing it
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 1, 1);
                var s2 = page1.DrawRectangle(1, 1, 2, 2);
                var s3 = page1.DrawRectangle(2, 2, 3, 3);
                MUT.Assert.AreEqual(3, page1.Shapes.Count);

                MUT.Assert.AreEqual(3, page1.Shapes.Count);
                undoscope0.Commit = false;

            }
            MUT.Assert.AreEqual(0, page1.Shapes.Count);

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Application_UndoScope_AbortNested()
        {

            var page1 = this.GetNewPage();
            var app = page1.Application;

            MUT.Assert.AreEqual(0, page1.Shapes.Count);

            // create a shape without undoing it
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 1, 1);
                MUT.Assert.AreEqual(1, page1.Shapes.Count);
                using (var undoscope1 = new VA.Application.UndoScope(app, "UndoScope2"))
                {
                    var s2 = page1.DrawRectangle(1, 1, 2, 2);
                    var s3 = page1.DrawRectangle(2, 2, 3, 3);
                    MUT.Assert.AreEqual(3, page1.Shapes.Count);

                    undoscope1.Commit = false;
                }

            }
            MUT.Assert.AreEqual(1, page1.Shapes.Count);

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Application_UndoScope_AbortOuter()
        {

            var page1 = this.GetNewPage();
            var app = page1.Application;

            MUT.Assert.AreEqual(0, page1.Shapes.Count);

            // create a shape without undoing it
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 1, 1);
                MUT.Assert.AreEqual(1, page1.Shapes.Count);
                using (var undoscope1 = new VA.Application.UndoScope(app, "UndoScope2"))
                {
                    var s2 = page1.DrawRectangle(1, 1, 2, 2);
                    var s3 = page1.DrawRectangle(2, 2, 3, 3);
                    MUT.Assert.AreEqual(3, page1.Shapes.Count);

                }
                MUT.Assert.AreEqual(3, page1.Shapes.Count);

                undoscope0.Commit = false;
            }
            MUT.Assert.AreEqual(0, page1.Shapes.Count);

            page1.Delete(0);
        }
    }
}