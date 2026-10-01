using System.Linq;
using VisioPowerShell.Commands.VisioApplication;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.PowerShell
{
    // Second slice of #173: parameter binding for Get-VisioDocument, Format-VisioWindow and the
    // pipeline-driven parameter sets of Out-VisioApplication.
    [MUT.TestClass]
    public class DocumentWindowCmdletBindingTests
    {
        private static readonly VisioPSSession Session = new VisioPSSession();

        [MUT.ClassInitialize]
        public static void ClassInitialize(MUT.TestContext context)
        {
            var new_visio_application = new NewVisioApplication();

            // Out-VisioApplication's input types are created with New-Object in the scripts below, which only
            // finds assemblies that are already loaded.
            var models_assembly = typeof(VisioAutomation.Models.Data.DataTableModel).Assembly;
        }

        [MUT.ClassCleanup]
        public static void ClassCleanup()
        {
            try { DocumentWindowCmdletBindingTests.Session.Cmd_Close_VisioApplication(true); }
            catch (System.Exception) { }
            DocumentWindowCmdletBindingTests.Session.CleanUp();
        }

        // -- Get-VisioDocument: parameter sets -----------------------------------------

        [MUT.TestMethod]
        public void GetVisioDocument_ActiveDocument_ReturnsTheActiveDocument()
        {
            var same = DocumentWindowCmdletBindingTests.Session.RunInNewDocument<bool>(
                "(Get-VisioDocument -ActiveDocument).Name -eq $doc.Name");
            MUT.Assert.IsTrue(same.Single());
        }

        [MUT.TestMethod]
        public void GetVisioDocument_NoArguments_IncludesTheNewDocument()
        {
            var found = DocumentWindowCmdletBindingTests.Session.RunInNewDocument<bool>(
                "(Get-VisioDocument | ForEach-Object { $_.Name }) -contains $doc.Name");
            MUT.Assert.IsTrue(found.Single());
        }

        [MUT.TestMethod]
        public void GetVisioDocument_PositionalName_FiltersByName()
        {
            var names = DocumentWindowCmdletBindingTests.Session.RunInNewDocument<string>(
                "(Get-VisioDocument $doc.Name).Name");
            MUT.Assert.AreEqual(1, names.Count);
        }

        [MUT.TestMethod]
        public void GetVisioDocument_ActiveDocumentWithName_FailsToResolveAParameterSet()
        {
            var ex = DocumentWindowCmdletBindingTests.Session.ExpectFailureInNewDocument("Get-VisioDocument -ActiveDocument -Name 'x'");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Parameter set cannot be resolved");
        }

        // -- Format-VisioWindow: three parameter sets ----------------------------------

        [MUT.TestMethod]
        public void FormatVisioWindow_Zoom_SetsTheWindowZoom()
        {
            var zoom = DocumentWindowCmdletBindingTests.Session.RunInNewDocument<double>(
                "Format-VisioWindow -Zoom 0.5; (Get-VisioApplication).ActiveWindow.Zoom");
            MUT.Assert.AreEqual(0.5, zoom.Single(), 1e-9);
        }

        [MUT.TestMethod]
        public void FormatVisioWindow_ZoomAndZoomTo_FailToResolveAParameterSet()
        {
            var ex = DocumentWindowCmdletBindingTests.Session.ExpectFailureInNewDocument("Format-VisioWindow -Zoom 1 -ZoomTo Page");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Parameter set cannot be resolved");
        }

        [MUT.TestMethod]
        public void FormatVisioWindow_UnknownZoomTo_FailsToBind()
        {
            var ex = DocumentWindowCmdletBindingTests.Session.ExpectFailureInNewDocument("Format-VisioWindow -ZoomTo Bogus");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Cannot bind parameter 'ZoomTo'");
        }

        // -- Out-VisioApplication: the parameter set is chosen by the type of the piped object ----

        [MUT.TestMethod]
        public void OutVisioApplication_PipedDataTableModel_DrawsTheTable()
        {
            var counts = DocumentWindowCmdletBindingTests.Session.RunInNewDocument<int>(
                "$dt = New-Object System.Data.DataTable; $null = $dt.Columns.Add('A'); $null = $dt.Columns.Add('B'); $null = $dt.Rows.Add('1', '2'); " +
                "$m = New-Object VisioAutomation.Models.Data.DataTableModel; $m.DataTable = $dt; " +
                "$m | Out-VisioApplication; (Get-VisioShape | Measure-Object).Count");
            MUT.Assert.AreEqual(2, counts.Single(), "one row of two columns is two cell shapes");
        }

        [MUT.TestMethod]
        public void OutVisioApplication_PipedXmlModel_DrawsTheTree()
        {
            var counts = DocumentWindowCmdletBindingTests.Session.RunInNewDocument<int>(
                "$m = New-Object VisioAutomation.Models.Data.XmlModel; $m.XmlDocument = [xml]'<root><a/><b/></root>'; " +
                "$m | Out-VisioApplication; (Get-VisioShape | Measure-Object).Count");
            MUT.Assert.AreEqual(5, counts.Single(), "three element nodes and two connectors");
        }

        [MUT.TestMethod]
        public void OutVisioApplication_PipedStringIsNotAModel_FailsToBind()
        {
            var ex = DocumentWindowCmdletBindingTests.Session.ExpectFailureInNewDocument("'not a model' | Out-VisioApplication");
            string message = CmdletScriptExtensions.MessageOf(ex);
            MUT.Assert.IsTrue(
                message.Contains("Parameter set cannot be resolved") || message.Contains("cannot be bound"),
                "expected a binding error, got: " + message);
        }
    }
}
