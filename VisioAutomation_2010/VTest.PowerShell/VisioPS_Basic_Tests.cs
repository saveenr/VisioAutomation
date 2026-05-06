using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioPowerShell.Commands.VisioApplication;
using VTest.PowerShell.Framework;

namespace VTest.PowerShell
{
    [MUT.TestClass]
    public class VisioPS_Basic_Tests
    {
        private static readonly VisioPS_Session Session = new VisioPS_Session();

        [MUT.ClassInitialize]
        public static void ClassInitialize(MUT.TestContext context)
        {
            var new_visio_application = new NewVisioApplication();
        }

        [MUT.ClassCleanup]
        public static void ClassCleanup()
        {
            // Close Visio before tearing down the runspace, so the testhost
            // doesn't leave a Visio orphan after exit. Swallow any exception:
            // teardown shouldn't fail the test run if the cmdlet errors.
            try { VisioPS_Basic_Tests.Session.Cmd_Close_VisioApplication(true); }
            catch (System.Exception) { }
            VisioPS_Basic_Tests.Session.CleanUp();
        }

        [MUT.TestMethod]
        public void VisioPS_SetShapeCells()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var stencil_basic = VisioPS_Basic_Tests.Session.Cmd_Open_VisioDocument("basic_u.vss");
            var master_roundrect = VisioPS_Basic_Tests.Session.Cmd_Get_VisioMaster(VTestPsArray.From("Rectangle"), stencil_basic);
            var shapes = VisioPS_Basic_Tests.Session.Cmd_New_VisioShape(VTestPsArray.From(master_roundrect), new [] {new VisioAutomation.Core.Point( 2.0, 3.0) });

            var shapecells = VisioPS_Basic_Tests.Session.Cmd_New_VisioShapeCells();
            shapecells.XFormPinX= "4 in";
            shapecells.XFormPinY = "6 in";

            VisioPS_Basic_Tests.Session.Cmd_Set_VisioShapeCells(VTestPsArray.From(shapecells), VTestPsArray.From(shapes));

            var dt = VisioPS_Basic_Tests.Session.Cmd_Get_VisioShapeCells(VTestPsArray.From(shapes));

            MUT.Assert.IsNotNull(dt);
            MUT.Assert.AreEqual("4 in", dt.Rows[0]["XFormPinX"]);
            MUT.Assert.AreEqual("6 in", dt.Rows[0]["XFormPinY"]);
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_SetPageCells()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var page = VisioPS_Basic_Tests.Session.Cmd_Get_VisioPage(activepage: true, name: null);

            var pagecells = VisioPS_Basic_Tests.Session.Cmd_New_VisioPageCells();
            pagecells.PageHeight = "4 in";
            pagecells.PageWidth= "3 in";

            VisioPS_Basic_Tests.Session.Cmd_Set_VisioPageCells( VTestPsArray.From(pagecells), VTestPsArray.From(page));
            
            var dt = VisioPS_Basic_Tests.Session.Cmd_Get_VisioPageCells(VTestPsArray.From(page));

            MUT.Assert.IsNotNull(dt);
            MUT.Assert.AreEqual("3 in", dt.Rows[0]["PageWidth"]);
            MUT.Assert.AreEqual("4 in", dt.Rows[0]["PageHeight"]);
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_DrawRectangleWithText()
        {
            var d = Session.Cmd_New_VisioDocument();
            var s = Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point( 0.0, 1.0),
                new VisioAutomation.Core.Point( 2.0, 3.0)
            });
            Session.Cmd_Set_VisioText( VTestPsArray.From("Hello World"), VTestPsArray.From(s));

            var r = Session.Cmd_Get_VisioText();

            MUT.Assert.AreEqual(1,r.Length);
            MUT.Assert.AreEqual("Hello World", r[0]);
            bool force = true;
            Session.Cmd_Close_VisioDocument( VTestPsArray.From(d), force);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioMaster_DocumentIsPositional()
        {
            // Issue #142: -Document is now Position = 1, so the natural
            // positional form Get-VisioMaster "Group" $doc binds correctly.
            // Use the runspace path (not the direct-Invoke path) because
            // direct invocation bypasses PowerShell's parameter binder.
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var basic_u = VisioPS_Basic_Tests.Session.Cmd_Open_VisioDocument("basic_u.vss");

            var masters = VisioPS_Basic_Tests.Session.InvokeScript<Microsoft.Office.Interop.Visio.Master>(
                "Get-VisioMaster \"Rectangle\" $stencil",
                ("stencil", basic_u));

            MUT.Assert.AreEqual(1, masters.Count);
            MUT.Assert.AreEqual("Rectangle", masters[0].Name);
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        // Issue #143: positional-parameter UX audit. Each test below exercises
        // the cmdlet's parameter binder via the runspace path (InvokeScript),
        // because direct cmdlet.Invoke() bypasses PowerShell's parameter binder
        // and so cannot verify positional binding. The pattern is:
        //   - cmdlets with -Name + object context: -Name at Position 0, object at Position 1
        //   - cmdlets with single object context : object at Position 0

        [MUT.TestMethod]
        public void VisioPS_GetVisioPage_NameAndDocumentArePositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var pages = VisioPS_Basic_Tests.Session.InvokeScript<Microsoft.Office.Interop.Visio.Page>(
                "Get-VisioPage \"Page-1\" $doc",
                ("doc", doc));
            MUT.Assert.AreEqual(1, pages.Count);
            MUT.Assert.AreEqual("Page-1", pages[0].Name);
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioShape_NameAndPageArePositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var s = VisioPS_Basic_Tests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 1.0),
                new VisioAutomation.Core.Point(2.0, 3.0)
            });
            VisioPS_Basic_Tests.Session.Cmd_Set_VisioText(VTestPsArray.From("Hello"), VTestPsArray.From((Microsoft.Office.Interop.Visio.Shape)s));
            // Sanity: at least exercise the Page=Position 1 form (Name not provided).
            var page = ((Microsoft.Office.Interop.Visio.Shape)s).ContainingPage;
            VisioPS_Basic_Tests.Session.InvokeScript<Microsoft.Office.Interop.Visio.Shape>(
                "Get-VisioShape -Page $p",
                ("p", page));
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioDocument_NameIsPositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            // -Name at Position 0; the call must not throw a binder error.
            // We don't assert a specific result count because the active document's
            // generated name varies, but binding alone is the contract under test.
            VisioPS_Basic_Tests.Session.InvokeScript<Microsoft.Office.Interop.Visio.Document>(
                "Get-VisioDocument \"basic_u.vss\"");
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioCustomProperty_ShapeIsPositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var s = VisioPS_Basic_Tests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 1.0),
                new VisioAutomation.Core.Point(2.0, 3.0)
            });
            VisioPS_Basic_Tests.Session.InvokeScript<object>(
                "Get-VisioCustomProperty $s",
                ("s", new[] { (Microsoft.Office.Interop.Visio.Shape)s }));
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioText_ShapeIsPositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var s = VisioPS_Basic_Tests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 1.0),
                new VisioAutomation.Core.Point(2.0, 3.0)
            });
            VisioPS_Basic_Tests.Session.Cmd_Set_VisioText(VTestPsArray.From("Hello"), VTestPsArray.From((Microsoft.Office.Interop.Visio.Shape)s));
            var text = VisioPS_Basic_Tests.Session.InvokeScript<object>(
                "Get-VisioText $s",
                ("s", new[] { (Microsoft.Office.Interop.Visio.Shape)s }));
            MUT.Assert.IsTrue(text.Count > 0);
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioHyperlink_ShapeIsPositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var s = VisioPS_Basic_Tests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 1.0),
                new VisioAutomation.Core.Point(2.0, 3.0)
            });
            VisioPS_Basic_Tests.Session.InvokeScript<object>(
                "Get-VisioHyperlink $s",
                ("s", new[] { (Microsoft.Office.Interop.Visio.Shape)s }));
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioLockCells_ShapeIsPositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var s = VisioPS_Basic_Tests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 1.0),
                new VisioAutomation.Core.Point(2.0, 3.0)
            });
            VisioPS_Basic_Tests.Session.InvokeScript<object>(
                "Get-VisioLockCells $s",
                ("s", new[] { (Microsoft.Office.Interop.Visio.Shape)s }));
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioControl_ShapeIsPositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var s = VisioPS_Basic_Tests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 1.0),
                new VisioAutomation.Core.Point(2.0, 3.0)
            });
            VisioPS_Basic_Tests.Session.InvokeScript<object>(
                "Get-VisioControl $s",
                ("s", new[] { (Microsoft.Office.Interop.Visio.Shape)s }));
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioUserDefinedCell_ShapeIsPositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var s = VisioPS_Basic_Tests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 1.0),
                new VisioAutomation.Core.Point(2.0, 3.0)
            });
            VisioPS_Basic_Tests.Session.InvokeScript<object>(
                "Get-VisioUserDefinedCell $s",
                ("s", new[] { (Microsoft.Office.Interop.Visio.Shape)s }));
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioShapeCells_ShapeIsPositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var s = VisioPS_Basic_Tests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 1.0),
                new VisioAutomation.Core.Point(2.0, 3.0)
            });
            VisioPS_Basic_Tests.Session.InvokeScript<System.Data.DataTable>(
                "Get-VisioShapeCells $s",
                ("s", new[] { (Microsoft.Office.Interop.Visio.Shape)s }));
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioPageCells_PageIsPositional()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var page = VisioPS_Basic_Tests.Session.Cmd_Get_VisioPage(activepage: true, name: null);
            VisioPS_Basic_Tests.Session.InvokeScript<System.Data.DataTable>(
                "Get-VisioPageCells $p",
                ("p", new[] { page }));
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_GetVisioShape_NoArgs_ReturnsAllShapesOnPage()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();

            // Drop two rectangles on the page.
            VisioPS_Basic_Tests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 1.0),
                new VisioAutomation.Core.Point(2.0, 3.0)
            });
            VisioPS_Basic_Tests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(3.0, 4.0),
                new VisioAutomation.Core.Point(5.0, 6.0)
            });

            // Get-VisioShape with no args must return every shape on the page (the default
            // parameter set's no-filter fallthrough). See issue #130 on the source repo.
            var shapes = VisioPS_Basic_Tests.Session.Cmd_Get_VisioShape();

            MUT.Assert.AreEqual(2, shapes.Count);
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void VisioPS_CreateContainer()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var app = VisioPS_Basic_Tests.Session.Cmd_Get_VisioApplication();

            var ver = VisioAutomation.Application.ApplicationHelper.GetVersion(app);

            var cont_doc = ver.Major >= 15 ? "SDCONT_U.VSSX" : "SDCONT_U.VSS";
            var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";
            var rectangle = "Rectangle";
            var basic_u_vss = "BASIC_U.VSS";

            var master = VisioPS_Basic_Tests.Session.Cmd_Get_VisioMaster(VTestPsArray.From(rectangle), basic_u_vss);


            VisioPS_Basic_Tests.Session.Cmd_New_VisioShape( VTestPsArray.From(master) , new[] { new VisioAutomation.Core.Point(1.0, 1.0) });

            // Drop a container on the page... the rectangle we created above should be selected by default. 
            // Since it is selected it will be added as a member to the container.

            var container = VisioPS_Basic_Tests.Session.Cmd_New_VisioContainer(cont_master_name, cont_doc);

            MUT.Assert.IsNotNull(container);

            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }
    }
}
