using System.Linq;
using VisioPowerShell.Commands.VisioApplication;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.PowerShell
{
    // Second slice of #173: parameter binding for the shape cmdlets (Get-VisioShape, Select-VisioShape,
    // Format-VisioShape, New-VisioHyperlink, Set-VisioCustomProperty, Connect-VisioShape).
    [MUT.TestClass]
    public class ShapeCmdletBindingTests
    {
        private static readonly VisioPSSession Session = new VisioPSSession();

        // Draws a 1 x 1 rectangle with its lower left corner at (x, y) on the active page.
        private static string rect(string variable, double x, double y)
        {
            return string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                "{0} = New-VisioShape -Rectangle -BoundingBox (New-VisioRectangle {1} {2} {3} {4}); ",
                variable, x, y, x + 1, y + 1);
        }

        [MUT.ClassInitialize]
        public static void ClassInitialize(MUT.TestContext context)
        {
            var new_visio_application = new NewVisioApplication();
        }

        [MUT.ClassCleanup]
        public static void ClassCleanup()
        {
            try { ShapeCmdletBindingTests.Session.Cmd_Close_VisioApplication(true); }
            catch (System.Exception) { }
            ShapeCmdletBindingTests.Session.CleanUp();
        }

        // -- Get-VisioShape: parameter sets ----------------------------------------

        [MUT.TestMethod]
        public void GetVisioShape_NoArguments_ReturnsEveryShapeOnThePage()
        {
            var counts = ShapeCmdletBindingTests.Session.RunInNewDocument<int>(
                rect("$a", 0, 0) + rect("$b", 3, 0) + "(Get-VisioShape | Measure-Object).Count");
            MUT.Assert.AreEqual(2, counts.Single());
        }

        [MUT.TestMethod]
        public void GetVisioShape_ID_ReturnsTheShapeWithThatShapeID()
        {
            var ok = ShapeCmdletBindingTests.Session.RunInNewDocument<bool>(
                rect("$a", 0, 0) + rect("$b", 3, 0) + "(Get-VisioShape -ID $b.ID).ID -eq $b.ID");
            MUT.Assert.IsTrue(ok.Single());
        }

        [MUT.TestMethod]
        public void GetVisioShape_ActiveSelectionWithName_FailsToResolveAParameterSet()
        {
            var ex = ShapeCmdletBindingTests.Session.ExpectFailureInNewDocument("Get-VisioShape -ActiveSelection -Name 'x'");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Parameter set cannot be resolved");
        }

        // -- Select-VisioShape: two required parameter sets --------------------------

        [MUT.TestMethod]
        public void SelectVisioShape_SelectAll_SelectsEveryShape()
        {
            var counts = ShapeCmdletBindingTests.Session.RunInNewDocument<int>(
                rect("$a", 0, 0) + rect("$b", 3, 0) +
                "Select-VisioShape -SelectionOperation SelectAll; (Get-VisioShape -ActiveSelection | Measure-Object).Count");
            MUT.Assert.AreEqual(2, counts.Single());
        }

        [MUT.TestMethod]
        public void SelectVisioShape_SelectNone_ClearsTheSelection()
        {
            var counts = ShapeCmdletBindingTests.Session.RunInNewDocument<int>(
                rect("$a", 0, 0) + "Select-VisioShape -SelectionOperation SelectAll; Select-VisioShape -SelectionOperation SelectNone; " +
                "(Get-VisioShape -ActiveSelection | Measure-Object).Count");
            MUT.Assert.AreEqual(0, counts.Single());
        }

        [MUT.TestMethod]
        public void SelectVisioShape_ShapesParameter_SelectsJustThoseShapes()
        {
            var counts = ShapeCmdletBindingTests.Session.RunInNewDocument<int>(
                rect("$a", 0, 0) + rect("$b", 3, 0) +
                "Select-VisioShape -Shapes $b; (Get-VisioShape -ActiveSelection | Measure-Object).Count");
            MUT.Assert.AreEqual(1, counts.Single());
        }

        [MUT.TestMethod]
        public void SelectVisioShape_WithNoArguments_FailsToResolveAParameterSet()
        {
            var ex = ShapeCmdletBindingTests.Session.ExpectFailureInNewDocument("Select-VisioShape");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Parameter set cannot be resolved");
        }

        [MUT.TestMethod]
        public void SelectVisioShape_ShapesAndSelectionOperationTogether_FailToResolveAParameterSet()
        {
            var ex = ShapeCmdletBindingTests.Session.ExpectFailureInNewDocument(
                rect("$a", 0, 0) + "Select-VisioShape -Shapes $a -SelectionOperation SelectAll");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Parameter set cannot be resolved");
        }

        [MUT.TestMethod]
        public void SelectVisioShape_UnknownOperation_FailsToBind()
        {
            var ex = ShapeCmdletBindingTests.Session.ExpectFailureInNewDocument("Select-VisioShape -SelectionOperation Bogus");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Cannot bind parameter 'SelectionOperation'");
        }

        // -- Format-VisioShape: numeric, switch and enum parameters -------------------

        [MUT.TestMethod]
        public void FormatVisioShape_NudgeX_MovesTheShapeByThatDistance()
        {
            var moved = ShapeCmdletBindingTests.Session.RunInNewDocument<double>(
                rect("$a", 0, 0) +
                "$x0 = $a.CellsU('PinX').ResultIU; Format-VisioShape -Shape $a -NudgeX 2; $a.CellsU('PinX').ResultIU - $x0");
            MUT.Assert.AreEqual(2.0, moved.Single(), 1e-9);
        }

        [MUT.TestMethod]
        public void FormatVisioShape_AlignHorizontalLeft_AlignsTheShapesLeftEdges()
        {
            var pins = ShapeCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$a", 0, 2) + rect("$b", 3, 2) +
                "Format-VisioShape -Shape $a,$b -AlignHorizontal Left; " +
                "\"$($a.CellsU('PinX').ResultIU),$($b.CellsU('PinX').ResultIU)\"");
            string[] parts = pins.Single().Split(',');
            MUT.Assert.AreEqual(double.Parse(parts[0]), double.Parse(parts[1]), 1e-9, "both shapes should end up in the same column");
        }

        [MUT.TestMethod]
        public void FormatVisioShape_DistributeHorizontalSwitch_SpacesTheShapesEvenly()
        {
            var pins = ShapeCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$a", 0, 0) + rect("$b", 1, 0) + rect("$c", 6, 0) +
                "Format-VisioShape -Shape $a,$b,$c -DistributeHorizontal; " +
                "\"$($a.CellsU('PinX').ResultIU),$($b.CellsU('PinX').ResultIU),$($c.CellsU('PinX').ResultIU)\"");
            double[] x = pins.Single().Split(',').Select(double.Parse).ToArray();
            MUT.Assert.AreEqual(x[1] - x[0], x[2] - x[1], 1e-9, "the gaps between the three shapes should be equal");
        }

        [MUT.TestMethod]
        public void FormatVisioShape_UnknownAlignment_FailsToBind()
        {
            var ex = ShapeCmdletBindingTests.Session.ExpectFailureInNewDocument("Format-VisioShape -AlignHorizontal Diagonal");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Cannot bind parameter 'AlignHorizontal'");
        }

        // -- New-VisioHyperlink: mandatory Address and boolean parameters ------------

        [MUT.TestMethod]
        public void NewVisioHyperlink_Address_AddsAHyperlinkToTheShape()
        {
            var results = ShapeCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "New-VisioHyperlink -Address 'https://example.com' -Description 'the site' -Shape $s; " +
                "\"$($s.Hyperlinks.Count)|$($s.Hyperlinks.Item(0).Address)|$($s.Hyperlinks.Item(0).Description)\"");
            MUT.Assert.AreEqual("1|https://example.com|the site", results.Single());
        }

        [MUT.TestMethod]
        public void NewVisioHyperlink_WithoutAnAddress_ReportsTheMissingMandatoryParameter()
        {
            var ex = ShapeCmdletBindingTests.Session.ExpectFailureInNewDocument(rect("$s", 0, 0) + "New-VisioHyperlink -Shape $s");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Address");
        }

        [MUT.TestMethod]
        public void NewVisioHyperlink_NewWindowDefaultAndInvisible_SetTheHyperlinkCells()
        {
            // Regression: these parameters bound but were dropped when the hyperlink was added,
            // so the cells always stayed FALSE.
            var results = ShapeCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "New-VisioHyperlink -Address 'https://example.com' -NewWindow $true -Default $true -Invisible $true -Shape $s; " +
                "\"$($s.CellsU('Hyperlink.Row_1.NewWindow').FormulaU),$($s.CellsU('Hyperlink.Row_1.Default').FormulaU),$($s.CellsU('Hyperlink.Row_1.Invisible').FormulaU)\"");
            MUT.Assert.AreEqual("TRUE,TRUE,TRUE", results.Single());
        }

        [MUT.TestMethod]
        public void NewVisioHyperlink_WithoutTheBooleanParameters_LeavesTheCellsFalse()
        {
            var results = ShapeCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "New-VisioHyperlink -Address 'https://example.com' -Shape $s; " +
                "\"$($s.CellsU('Hyperlink.Row_1.NewWindow').FormulaU),$($s.CellsU('Hyperlink.Row_1.Default').FormulaU),$($s.CellsU('Hyperlink.Row_1.Invisible').FormulaU)\"");
            MUT.Assert.AreEqual("FALSE,FALSE,FALSE", results.Single());
        }

        [MUT.TestMethod]
        public void NewVisioHyperlink_SortKey_SetsTheSortKeyCell()
        {
            var results = ShapeCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "New-VisioHyperlink -Address 'https://example.com' -SortKey 'b-key' -Shape $s; " +
                "$s.CellsU('Hyperlink.Row_1.SortKey').ResultStr(0)");
            MUT.Assert.AreEqual("b-key", results.Single());
        }

        // -- Set-VisioCustomProperty: two parameter sets -------------------------------

        [MUT.TestMethod]
        public void SetVisioCustomProperty_NamedProperties_SetValueTypeAndLabel()
        {
            var results = ShapeCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "Set-VisioCustomProperty 'cost' 2.5 -Type 2 -Label 'Cost' -Shape $s; " +
                "\"$($s.CellsU('Prop.cost').FormulaU)|$($s.CellsU('Prop.cost.Type').ResultIU)|$($s.CellsU('Prop.cost.Label').FormulaU)\"");
            MUT.Assert.AreEqual("2.5|2|\"Cost\"", results.Single());
        }

        [MUT.TestMethod]
        public void SetVisioCustomProperty_CellsAndValueTogether_FailToResolveAParameterSet()
        {
            var ex = ShapeCmdletBindingTests.Session.ExpectFailureInNewDocument(
                rect("$s", 0, 0) +
                "$c = New-Object VisioAutomation.Shapes.CustomPropertyCells; Set-VisioCustomProperty 'p' -Cells $c -Value 1 -Shape $s");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Parameter set cannot be resolved");
        }

        [MUT.TestMethod]
        public void SetVisioCustomProperty_WithoutAName_FailsToBind()
        {
            var ex = ShapeCmdletBindingTests.Session.ExpectFailureInNewDocument("Set-VisioCustomProperty");
            MUT.Assert.IsTrue(CmdletScriptExtensions.MessageOf(ex).Contains("Parameter set cannot be resolved") ||
                              CmdletScriptExtensions.MessageOf(ex).Contains("missing mandatory"),
                              "expected a binding error, got: " + CmdletScriptExtensions.MessageOf(ex));
        }

        // -- Connect-VisioShape: mandatory From and To ---------------------------------

        [MUT.TestMethod]
        public void ConnectVisioShape_WithoutFromAndTo_ReportsTheMissingMandatoryParameters()
        {
            var ex = ShapeCmdletBindingTests.Session.ExpectFailureInNewDocument("Connect-VisioShape");
            string message = CmdletScriptExtensions.MessageOf(ex);
            MUT.StringAssert.Contains(message, "From");
            MUT.StringAssert.Contains(message, "To");
        }
    }
}
