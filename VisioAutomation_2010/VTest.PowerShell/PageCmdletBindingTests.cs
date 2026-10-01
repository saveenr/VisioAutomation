using System.Linq;
using VisioPowerShell.Commands.VisioApplication;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.PowerShell
{
    // Second slice of #173: parameter binding for the page cmdlets (Get-VisioPage, New-VisioPage,
    // Format-VisioPage, Remove-VisioPage). Each test runs a script in the session's runspace so
    // PowerShell's binder resolves parameter sets, positions, switches and enums.
    [MUT.TestClass]
    public class PageCmdletBindingTests
    {
        private static readonly VisioPSSession Session = new VisioPSSession();

        [MUT.ClassInitialize]
        public static void ClassInitialize(MUT.TestContext context)
        {
            var new_visio_application = new NewVisioApplication();
        }

        [MUT.ClassCleanup]
        public static void ClassCleanup()
        {
            try { PageCmdletBindingTests.Session.Cmd_Close_VisioApplication(true); }
            catch (System.Exception) { }
            PageCmdletBindingTests.Session.CleanUp();
        }

        private static string page_size(string page_expression)
        {
            return "\"$(" + page_expression + ".PageSheet.CellsU('PageWidth').ResultIU) x $(" + page_expression + ".PageSheet.CellsU('PageHeight').ResultIU)\"";
        }

        // -- Get-VisioPage: parameter sets -----------------------------------------

        [MUT.TestMethod]
        public void GetVisioPage_NoArguments_ReturnsEveryPageInTheDocument()
        {
            var counts = PageCmdletBindingTests.Session.RunInNewDocument<int>(
                "$null = New-VisioPage -Name 'second'; (Get-VisioPage | Measure-Object).Count");
            MUT.Assert.AreEqual(2, counts.Single());
        }

        [MUT.TestMethod]
        public void GetVisioPage_PositionalName_FiltersByName()
        {
            var names = PageCmdletBindingTests.Session.RunInNewDocument<string>(
                "$null = New-VisioPage -Name 'second'; (Get-VisioPage 'second').Name");
            MUT.Assert.AreEqual("second", names.Single());
        }

        [MUT.TestMethod]
        public void GetVisioPage_ActivePage_ReturnsTheActivePage()
        {
            var same = PageCmdletBindingTests.Session.RunInNewDocument<bool>(
                "$null = New-VisioPage -Name 'second'; " +
                "(Get-VisioPage -ActivePage).Name -eq (Get-VisioApplication).ActivePage.Name");
            MUT.Assert.IsTrue(same.Single());
        }

        [MUT.TestMethod]
        public void GetVisioPage_ActivePageWithName_FailsToResolveAParameterSet()
        {
            var ex = PageCmdletBindingTests.Session.ExpectFailureInNewDocument("Get-VisioPage -ActivePage -Name 'x'");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Parameter set cannot be resolved");
        }

        // -- New-VisioPage ----------------------------------------------------------

        [MUT.TestMethod]
        public void NewVisioPage_Name_SetsThePageName()
        {
            var names = PageCmdletBindingTests.Session.RunInNewDocument<string>("(New-VisioPage -Name 'Overview').Name");
            MUT.Assert.AreEqual("Overview", names.Single());
        }

        [MUT.TestMethod]
        public void NewVisioPage_EmptyName_Throws()
        {
            var ex = PageCmdletBindingTests.Session.ExpectFailureInNewDocument("New-VisioPage -Name ''");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Name can't be empty");
        }

        [MUT.TestMethod]
        public void NewVisioPage_WhitespaceName_Throws()
        {
            var ex = PageCmdletBindingTests.Session.ExpectFailureInNewDocument("New-VisioPage -Name '   '");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Name can't be empty");
        }

        [MUT.TestMethod]
        public void NewVisioPage_WidthAndHeight_SetThePageSize()
        {
            var sizes = PageCmdletBindingTests.Session.RunInNewDocument<string>(
                "$p = New-VisioPage -Width 3 -Height 2; " + page_size("$p"));
            MUT.Assert.AreEqual("3 x 2", sizes.Single());
        }

        // -- Format-VisioPage -------------------------------------------------------

        [MUT.TestMethod]
        public void FormatVisioPage_WidthAndHeight_SetThePageSize()
        {
            // Regression: SetFormatCells committed the writes to the Page instead of its page sheet,
            // so Format-VisioPage -Width / -Height always threw a COMException.
            var sizes = PageCmdletBindingTests.Session.RunInNewDocument<string>(
                "$p = New-VisioPage -Name 'f'; Format-VisioPage -Page $p -Width 5 -Height 4; " + page_size("$p"));
            MUT.Assert.AreEqual("5 x 4", sizes.Single());
        }

        [MUT.TestMethod]
        public void FormatVisioPage_OnlyWidth_LeavesTheHeightAlone()
        {
            var sizes = PageCmdletBindingTests.Session.RunInNewDocument<string>(
                "$p = New-VisioPage -Width 3 -Height 2; Format-VisioPage -Page $p -Width 6; " + page_size("$p"));
            MUT.Assert.AreEqual("6 x 2", sizes.Single());
        }

        [MUT.TestMethod]
        public void FormatVisioPage_Landscape_MakesThePageWiderThanItIsTall()
        {
            var wider = PageCmdletBindingTests.Session.RunInNewDocument<bool>(
                "$p = New-VisioPage -Width 4 -Height 6; Format-VisioPage -Page $p -Orientation Landscape; " +
                "$p.PageSheet.CellsU('PageWidth').ResultIU -gt $p.PageSheet.CellsU('PageHeight').ResultIU");
            MUT.Assert.IsTrue(wider.Single());
        }

        [MUT.TestMethod]
        public void FormatVisioPage_UnknownOrientation_FailsToBind()
        {
            var ex = PageCmdletBindingTests.Session.ExpectFailureInNewDocument("Format-VisioPage -Orientation Sideways");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Cannot bind parameter 'Orientation'");
        }

        [MUT.TestMethod]
        public void FormatVisioPage_FitContents_ShrinksThePageAroundTheShapes()
        {
            var widths = PageCmdletBindingTests.Session.RunInNewDocument<double>(
                "$p = New-VisioPage -Width 20 -Height 20; " +
                "$null = New-VisioShape -Rectangle -BoundingBox (New-VisioRectangle 2 2 3 3); " +
                "Format-VisioPage -Page $p -FitContents -BorderWidth 0.5 -BorderHeight 0.5; " +
                "$p.PageSheet.CellsU('PageWidth').ResultIU");
            MUT.Assert.IsTrue(widths.Single() < 20.0, "the page should shrink to fit its contents, got " + widths.Single());
            MUT.Assert.IsTrue(widths.Single() >= 1.0, "the page should still hold the shape, got " + widths.Single());
        }

        // -- Remove-VisioPage -------------------------------------------------------

        [MUT.TestMethod]
        public void RemoveVisioPage_PipedPage_RemovesThatPage()
        {
            var counts = PageCmdletBindingTests.Session.RunInNewDocument<string>(
                "$p = New-VisioPage -Name 'gone'; $before = (Get-VisioPage | Measure-Object).Count; " +
                "$p | Remove-VisioPage; $after = (Get-VisioPage | Measure-Object).Count; " +
                "\"$before -> $after\"");
            MUT.Assert.AreEqual("2 -> 1", counts.Single());
        }
    }
}
