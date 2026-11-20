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

        // -- New-VisioDocument: -Template creates the drawing from the template (#229) ---------------

        [MUT.TestMethod]
        public void NewVisioDocument_Template_CreatesTheDrawingFromThatTemplate()
        {
            var results = DocumentWindowCmdletBindingTests.Session.InvokeScriptStrict<string>(
                "$ver = (Get-VisioApplication).Version; " +
                "$name = if ([int]($ver -split '[.,]')[0] -ge 15) { 'basflo_u.vstx' } else { 'basflo_u.vst' }; " +
                "$d = New-VisioDocument -Template $name; " +
                "try { $d.Template } finally { try { $d.Saved = $true; $d.Close() } catch { } }");
            MUT.Assert.IsTrue(results.Single().ToLowerInvariant().Contains("basflo_u"), "the new drawing should be based on the template: '" + results.Single() + "'");
        }

        [MUT.TestMethod]
        public void NewVisioDocument_TemplateThatIsAStencil_FailsWithAClearMessage()
        {
            var ex = DocumentWindowCmdletBindingTests.Session.ExpectFailureInNewDocument("New-VisioDocument -Template 'basic_u.vss'");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "is a stencil, not a template");
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
        public void OutVisioApplication_PipedOrgChart_WithNoDocumentOpen_DrawsTheChartInANewDocument()
        {
            // The org chart is drawn in a new document, so it must not need a document or page to already be open.
            // This does not use RunInNewDocument, which would open a document first.
            var results = DocumentWindowCmdletBindingTests.Session.InvokeScriptStrict<int>(
                "$app = Get-VisioApplication; " +
                "foreach ($d in @($app.Documents)) { $d.Saved = $true; $d.Close() }; " +
                "try { " +
                "$drawings = { @($app.Documents | Where-Object { $_.Type -eq 1 }).Count }; " + // 1 = a drawing; the Documents collection also holds stencils
                "$before = & $drawings; " +
                "$chart = New-Object VisioAutomation.Models.Documents.OrgCharts.OrgChartDocument; " +
                "$node = New-Object VisioAutomation.Models.Documents.OrgCharts.Node -ArgumentList 'A'; " +
                "$chart.OrgCharts.Add($node); " +
                "$chart | Out-VisioApplication; " +
                "$before; & $drawings; (Get-VisioShape | Measure-Object).Count " +
                "} finally { foreach ($d in @($app.Documents)) { try { $d.Saved = $true; $d.Close() } catch { } } }");

            MUT.Assert.AreEqual(3, results.Count, "the script writes the drawing count before, the drawing count after, and the shape count");
            MUT.Assert.AreEqual(0, results[0], "the test needs Visio attached with no drawing open");
            MUT.Assert.AreEqual(1, results[1], "the chart should be drawn in one new drawing");
            MUT.Assert.IsTrue(results[2] > 0, "the new document should contain the chart's shapes");
        }

        // Draws a two-node directed graph with Out-VisioApplication and returns how many shapes were made from the
        // given master. "styling_setup" runs before the draw; "out_arguments" is added to the Out-VisioApplication call.
        // Like the org chart test, this closes every document first and afterward, because the graph is drawn in a new document.
        private static int count_shapes_from_master_after_drawing_a_graph(string styling_setup, string out_arguments, string master_name)
        {
            var results = DocumentWindowCmdletBindingTests.Session.InvokeScriptStrict<int>(
                "$app = Get-VisioApplication; " +
                "foreach ($d in @($app.Documents)) { $d.Saved = $true; $d.Close() }; " +
                "try { " +
                "$layout = New-Object VisioAutomation.Models.Layouts.DirectedGraph.DirectedGraphLayout; " +
                "$a = $layout.AddNode('n0', 'A', 'basic_u.vss', 'Rectangle'); " +
                "$b = $layout.AddNode('n1', 'B', 'basic_u.vss', 'Rectangle'); " +
                "$null = $layout.AddEdge('c0', $a, $b, '', [VisioAutomation.Models.ConnectorType]::Curved); " +
                "$dgdoc = New-Object VisioAutomation.Models.Layouts.DirectedGraph.DirectedGraphDocument; " +
                "$dgdoc.Layouts.Add($layout); " +
                styling_setup + " " +
                "$dgdoc | Out-VisioApplication " + out_arguments + "; " +
                "(Get-VisioShape | Where-Object { $_.Master -ne $null -and $_.Master.NameU -eq '" + master_name + "' } | Measure-Object).Count " +
                "} finally { foreach ($d in @($app.Documents)) { try { $d.Saved = $true; $d.Close() } catch { } } }");
            return results.Single();
        }

        [MUT.TestMethod]
        public void OutVisioApplication_DirectedGraphWithoutStyling_UsesTheDefaultConnector()
        {
            int count = DocumentWindowCmdletBindingTests.count_shapes_from_master_after_drawing_a_graph("", "", "Dynamic Connector");
            MUT.Assert.AreEqual(1, count, "the one edge should use the default Dynamic Connector master");
        }

        [MUT.TestMethod]
        public void OutVisioApplication_DirectedGraphStyling_ChangesTheConnectorMaster()
        {
            int count = DocumentWindowCmdletBindingTests.count_shapes_from_master_after_drawing_a_graph(
                "$styling = New-Object VisioAutomation.Models.Layouts.DirectedGraph.DirectedGraphStyling; $styling.EdgeMasterName = 'Line-curve connector';",
                "-DirectedGraphStyling $styling",
                "Line-curve connector");
            MUT.Assert.AreEqual(1, count, "the one edge should use the connector master named in the styling");
        }

        [MUT.TestMethod]
        public void OutVisioApplication_DirectedGraphStylingWithAnOrgChart_FailsToBind()
        {
            // -DirectedGraphStyling belongs to the directed graph parameter set only.
            var ex = DocumentWindowCmdletBindingTests.Session.ExpectFailureInNewDocument(
                "$styling = New-Object VisioAutomation.Models.Layouts.DirectedGraph.DirectedGraphStyling; " +
                "$chart = New-Object VisioAutomation.Models.Documents.OrgCharts.OrgChartDocument; " +
                "$chart | Out-VisioApplication -DirectedGraphStyling $styling");
            string message = CmdletScriptExtensions.MessageOf(ex);
            MUT.Assert.IsTrue(
                message.Contains("Parameter set cannot be resolved") || message.Contains("cannot be bound"),
                "expected a binding error, got: " + message);
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
