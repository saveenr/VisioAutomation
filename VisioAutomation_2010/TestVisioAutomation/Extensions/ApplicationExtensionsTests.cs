using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class ApplicationExtensionsTests : VisioAutomationTest
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


        [TestMethod]
        public void TestUndoScopeScenarios()
        {
            this.AbortNestedUndoScope();
            this.AbortOuterUndoScope();
            this.AbortUndoScope();
            this.SimpleUndoScope();
            this.UndoScopeNestedInner();
            this.UndoScopeNestedOuter();
        }

        public void SimpleUndoScope()
        {
            var page1 = GetNewPage();
            var app = page1.Application;

            Assert.AreEqual(0, page1.Shapes.Count);

            // create a shape without undoing it
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 2, 2);
                Assert.AreEqual(1, page1.Shapes.Count);
            }
            Assert.AreEqual(1, page1.Shapes.Count);

            // create a shape and undo it
            using (var undoscope1 = new VA.Application.UndoScope(app, "UndoScope2"))
            {
                var s1 = page1.DrawRectangle(1, 1, 3, 3);
                Assert.AreEqual(2, page1.Shapes.Count);
            }
            app.Undo();

            Assert.AreEqual(1, page1.Shapes.Count);
            page1.Delete(0);
        }

        public void UndoScopeNestedInner()
        {
            var page1 = GetNewPage();
            var app = page1.Application;

            Assert.AreEqual(0, page1.Shapes.Count);

            // Test that inner undo doesn't affect outer
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 2, 2);
                Assert.AreEqual(1, page1.Shapes.Count);

                // create a shape and undo it
                using (var undoscope1 = new VA.Application.UndoScope(app, "UndoScope2"))
                {
                    var s2 = page1.DrawRectangle(1, 1, 3, 3);
                    Assert.AreEqual(2, page1.Shapes.Count);
                }
                app.Undo();
            }
            Assert.AreEqual(1, page1.Shapes.Count);
            page1.Delete(0);
        }

        public void UndoScopeNestedOuter()
        {
            var page1 = GetNewPage();
            var app = page1.Application;

            Assert.AreEqual(0, page1.Shapes.Count);

            // Test that outter does affect inner
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 2, 2);
                Assert.AreEqual(1, page1.Shapes.Count);

                // create a shape and undo it
                using (var undoscope1 = new VA.Application.UndoScope(app, "UndoScope2"))
                {
                    var s2 = page1.DrawRectangle(1, 1, 3, 3);
                    Assert.AreEqual(2, page1.Shapes.Count);
                }
            }
            app.Undo();

            Assert.AreEqual(0, page1.Shapes.Count);
            page1.Delete(0);
        }


        public void AbortUndoScope()
        {

            var page1 = GetNewPage();
            var app = page1.Application;

            Assert.AreEqual(0, page1.Shapes.Count);

            // create a shape without undoing it
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 1, 1);
                var s2 = page1.DrawRectangle(1, 1, 2, 2);
                var s3 = page1.DrawRectangle(2, 2, 3, 3);
                Assert.AreEqual(3, page1.Shapes.Count);

                Assert.AreEqual(3, page1.Shapes.Count);
                undoscope0.Commit = false;

            }
            Assert.AreEqual(0, page1.Shapes.Count);

            page1.Delete(0);
        }

        public void AbortNestedUndoScope()
        {

            var page1 = GetNewPage();
            var app = page1.Application;

            Assert.AreEqual(0, page1.Shapes.Count);

            // create a shape without undoing it
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 1, 1);
                Assert.AreEqual(1, page1.Shapes.Count);
                using (var undoscope1 = new VA.Application.UndoScope(app, "UndoScope2"))
                {
                    var s2 = page1.DrawRectangle(1, 1, 2, 2);
                    var s3 = page1.DrawRectangle(2, 2, 3, 3);
                    Assert.AreEqual(3, page1.Shapes.Count);

                    undoscope1.Commit = false;
                }

            }
            Assert.AreEqual(1, page1.Shapes.Count);

            page1.Delete(0);
        }

        public void AbortOuterUndoScope()
        {

            var page1 = GetNewPage();
            var app = page1.Application;

            Assert.AreEqual(0, page1.Shapes.Count);

            // create a shape without undoing it
            using (var undoscope0 = new VA.Application.UndoScope(app, "UndoScope1"))
            {
                var s1 = page1.DrawRectangle(0, 0, 1, 1);
                Assert.AreEqual(1, page1.Shapes.Count);
                using (var undoscope1 = new VA.Application.UndoScope(app, "UndoScope2"))
                {
                    var s2 = page1.DrawRectangle(1, 1, 2, 2);
                    var s3 = page1.DrawRectangle(2, 2, 3, 3);
                    Assert.AreEqual(3, page1.Shapes.Count);

                }
                Assert.AreEqual(3, page1.Shapes.Count);

                undoscope0.Commit = false;
            }
            Assert.AreEqual(0, page1.Shapes.Count);

            page1.Delete(0);
        }

    }
}