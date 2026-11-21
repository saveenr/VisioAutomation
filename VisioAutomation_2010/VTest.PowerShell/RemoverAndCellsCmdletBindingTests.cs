using System.Linq;
using VisioPowerShell.Commands.VisioApplication;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.PowerShell
{
    // Third slice of #173 (tracked in #263), first group: the cmdlets that remove or overwrite things
    // (Remove-VisioShape, Remove-VisioCustomProperty, Remove-VisioHyperlink, Remove-VisioUserDefinedCell,
    // Remove-VisioControl, Set-VisioShapeCells, Set-VisioPageCells), plus the builders they need
    // (New-VisioControl, New-VisioShapeCells, New-VisioPageCells). These are the cmdlets most likely to
    // hide a binding bug that silently does nothing, which is the class of bug slice 2 found.
    // Every assertion reads the result back from Visio inside the script, so COM objects never cross
    // into the test.
    [MUT.TestClass]
    public class RemoverAndCellsCmdletBindingTests
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
            try { RemoverAndCellsCmdletBindingTests.Session.Cmd_Close_VisioApplication(true); }
            catch (System.Exception) { }
            RemoverAndCellsCmdletBindingTests.Session.CleanUp();
        }

        private static string MissingMandatory(System.Exception ex, string parameter)
        {
            var message = CmdletScriptExtensions.MessageOf(ex);
            MUT.StringAssert.Contains(message, parameter);
            return message;
        }

        // -- Remove-VisioShape -----------------------------------------------------

        [MUT.TestMethod]
        public void RemoveVisioShape_ShapeParameter_RemovesJustThatShape()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$a", 0, 0) + rect("$b", 3, 0) +
                "$aid = $a.ID; $bid = $b.ID; Remove-VisioShape -Shape $a; " +
                "$left = @(Get-VisioShape); \"$($left.Count)|$($left[0].ID -eq $bid)\"");
            MUT.Assert.AreEqual("1|True", results.Single());
        }

        [MUT.TestMethod]
        public void RemoveVisioShape_ShapeArray_RemovesEveryShapeInTheArray()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$a", 0, 0) + rect("$b", 3, 0) + rect("$c", 6, 0) +
                "$cid = $c.ID; Remove-VisioShape -Shape @($a, $b); " +
                "$left = @(Get-VisioShape); \"$($left.Count)|$($left[0].ID -eq $cid)\"");
            MUT.Assert.AreEqual("1|True", results.Single());
        }

        [MUT.TestMethod]
        public void RemoveVisioShape_WithoutShapeParameter_RemovesTheSelectedShapes()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$a", 0, 0) + rect("$b", 3, 0) +
                "$aid = $a.ID; Select-VisioShape -Shapes $b; Remove-VisioShape; " +
                "$left = @(Get-VisioShape); \"$($left.Count)|$($left[0].ID -eq $aid)\"");
            MUT.Assert.AreEqual("1|True", results.Single());
        }

        // -- Remove-VisioCustomProperty ----------------------------------------------

        [MUT.TestMethod]
        public void RemoveVisioCustomProperty_Name_RemovesThatPropertyOnly()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "Set-VisioCustomProperty 'cost' 2.5 -Shape $s; Set-VisioCustomProperty 'other' 1 -Shape $s; " +
                "Remove-VisioCustomProperty -Name 'cost' -Shape $s; " +
                "\"$($s.CellExistsU('Prop.cost', 0))|$($s.CellExistsU('Prop.other', 0))\"");
            MUT.Assert.AreEqual("0|-1", results.Single());
        }

        [MUT.TestMethod]
        public void RemoveVisioCustomProperty_NameIsPositional()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<int>(
                rect("$s", 0, 0) +
                "Set-VisioCustomProperty 'cost' 2.5 -Shape $s; Remove-VisioCustomProperty 'cost' -Shape $s; " +
                "$s.CellExistsU('Prop.cost', 0)");
            MUT.Assert.AreEqual(0, results.Single());
        }

        [MUT.TestMethod]
        public void RemoveVisioCustomProperty_WithoutAName_ReportsTheMissingMandatoryParameter()
        {
            var ex = RemoverAndCellsCmdletBindingTests.Session.ExpectFailureInNewDocument(
                rect("$s", 0, 0) + "Remove-VisioCustomProperty -Shape $s");
            MissingMandatory(ex, "Name");
        }

        // -- Remove-VisioHyperlink ---------------------------------------------------

        [MUT.TestMethod]
        public void RemoveVisioHyperlink_Index_RemovesTheHyperlinkAtThatIndex()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "New-VisioHyperlink -Address 'https://first.example' -Shape $s; " +
                "New-VisioHyperlink -Address 'https://second.example' -Shape $s; " +
                "Remove-VisioHyperlink -Index 0 -Shape $s; " +
                "\"$($s.Hyperlinks.Count)|$($s.Hyperlinks.Item(0).Address)\"");
            MUT.Assert.AreEqual("1|https://second.example", results.Single());
        }

        [MUT.TestMethod]
        public void RemoveVisioHyperlink_IndexIsPositional()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<int>(
                rect("$s", 0, 0) +
                "New-VisioHyperlink -Address 'https://first.example' -Shape $s; " +
                "Remove-VisioHyperlink 0 -Shape $s; $s.Hyperlinks.Count");
            MUT.Assert.AreEqual(0, results.Single());
        }

        [MUT.TestMethod]
        public void RemoveVisioHyperlink_WithoutAnIndex_ReportsTheMissingMandatoryParameter()
        {
            var ex = RemoverAndCellsCmdletBindingTests.Session.ExpectFailureInNewDocument(
                rect("$s", 0, 0) + "Remove-VisioHyperlink -Shape $s");
            MissingMandatory(ex, "Index");
        }

        [MUT.TestMethod]
        public void RemoveVisioHyperlink_NonNumericIndex_FailsToBind()
        {
            var ex = RemoverAndCellsCmdletBindingTests.Session.ExpectFailureInNewDocument(
                rect("$s", 0, 0) + "Remove-VisioHyperlink -Index 'first' -Shape $s");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Cannot bind parameter 'Index'");
        }

        // -- Remove-VisioUserDefinedCell ---------------------------------------------

        [MUT.TestMethod]
        public void RemoveVisioUserDefinedCell_Name_RemovesThatCellOnly()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "Set-VisioUserDefinedCell -Name 'one' -Value 'x' -Shape $s; Set-VisioUserDefinedCell -Name 'two' -Value 'y' -Shape $s; " +
                "Remove-VisioUserDefinedCell -Name 'one' -Shape $s; " +
                "\"$($s.CellExistsU('User.one', 0))|$($s.CellExistsU('User.two', 0))\"");
            MUT.Assert.AreEqual("0|-1", results.Single());
        }

        [MUT.TestMethod]
        public void RemoveVisioUserDefinedCell_WithoutAName_ReportsTheMissingMandatoryParameter()
        {
            var ex = RemoverAndCellsCmdletBindingTests.Session.ExpectFailureInNewDocument(
                rect("$s", 0, 0) + "Remove-VisioUserDefinedCell -Shape $s");
            MissingMandatory(ex, "Name");
        }

        // -- New-VisioControl and Remove-VisioControl --------------------------------

        [MUT.TestMethod]
        public void NewVisioControl_XAndY_AddsAControlWithThoseFormulas()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "New-VisioControl -X 'Width*0.5' -Y 'Height*0.25' -Shape $s; " +
                "\"$($s.CellExistsU('Controls.Row_1', 0))|$($s.CellsU('Controls.Row_1.X').FormulaU)|$($s.CellsU('Controls.Row_1.Y').FormulaU)\"");
            MUT.Assert.AreEqual("-1|Width*0.5|Height*0.25", results.Single());
        }

        [MUT.TestMethod]
        public void NewVisioControl_CanGlueTrue_SetsTheCanGlueCell()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "New-VisioControl -X 'Width*0.5' -Y 'Height*0.5' -CanGlue $true -Shape $s; " +
                "$s.CellsU('Controls.Row_1.CanGlue').FormulaU");
            MUT.Assert.AreEqual("TRUE", results.Single());
        }

        [MUT.TestMethod]
        public void NewVisioControl_WithoutCanGlue_LeavesTheCanGlueCellFalse()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "New-VisioControl -X 'Width*0.5' -Y 'Height*0.5' -Shape $s; " +
                "$s.CellsU('Controls.Row_1.CanGlue').FormulaU");
            MUT.Assert.AreEqual("FALSE", results.Single());
        }

        [MUT.TestMethod]
        public void RemoveVisioControl_Index_RemovesTheControlAtThatIndex()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$s", 0, 0) +
                "New-VisioControl -X 'Width*0.25' -Y 'Height*0.5' -Shape $s; " +
                "New-VisioControl -X 'Width*0.75' -Y 'Height*0.5' -Shape $s; " +
                "Remove-VisioControl -Index 0 -Shape $s; " +
                "\"$($s.CellExistsU('Controls.Row_1', 0))|$($s.CellExistsU('Controls.Row_2', 0))|$($s.CellsU('Controls.Row_2.X').FormulaU)\"");
            // Visio row names are stable: the surviving control keeps the name it was created with (Row_2).
            MUT.Assert.AreEqual("0|-1|Width*0.75", results.Single());
        }

        [MUT.TestMethod]
        public void RemoveVisioControl_WithoutAnIndex_ReportsTheMissingMandatoryParameter()
        {
            var ex = RemoverAndCellsCmdletBindingTests.Session.ExpectFailureInNewDocument(
                rect("$s", 0, 0) + "Remove-VisioControl -Shape $s");
            MissingMandatory(ex, "Index");
        }

        // -- New-VisioShapeCells and New-VisioPageCells -------------------------------

        [MUT.TestMethod]
        public void NewVisioShapeCells_NoCount_ReturnsOneObject()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                "$c = @(New-VisioShapeCells); \"$($c.Count)|$($c[0].GetType().Name)\"");
            MUT.Assert.AreEqual("1|ShapeCells", results.Single());
        }

        [MUT.TestMethod]
        public void NewVisioShapeCells_Count_ReturnsACollectionOfThatManyIndependentObjects()
        {
            // The collection comes back as one List object, not as separate pipeline objects, so do not
            // wrap the call in @(...) here (see #291 for the mismatch with the docs, which say "array").
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                "$c = New-VisioShapeCells -Count 3; $c[0].XFormWidth = '2'; " +
                "\"$($c.Count)|$($c[0].XFormWidth)|$($c[1].XFormWidth -eq $null)\"");
            MUT.Assert.AreEqual("3|2|True", results.Single());
        }

        [MUT.TestMethod]
        public void NewVisioShapeCells_CountZero_ReturnsAnEmptyCollection()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<int>(
                "$c = New-VisioShapeCells -Count 0; $c.Count");
            MUT.Assert.AreEqual(0, results.Single());
        }

        [MUT.TestMethod]
        public void NewVisioPageCells_NoCount_ReturnsOneObject()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                "$c = @(New-VisioPageCells); \"$($c.Count)|$($c[0].GetType().Name)\"");
            MUT.Assert.AreEqual("1|PageCells", results.Single());
        }

        [MUT.TestMethod]
        public void NewVisioPageCells_Count_ReturnsACollectionOfThatManyObjects()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<int>(
                "$c = New-VisioPageCells -Count 4; $c.Count");
            MUT.Assert.AreEqual(4, results.Single());
        }

        // -- Set-VisioShapeCells -------------------------------------------------------

        [MUT.TestMethod]
        public void SetVisioShapeCells_Cells_WritesTheFormulasToTheShape()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<double>(
                rect("$s", 0, 0) +
                "$c = New-VisioShapeCells; $c.XFormWidth = '3'; Set-VisioShapeCells -Cells $c -Shape $s; " +
                "$s.CellsU('Width').ResultIU");
            MUT.Assert.AreEqual(3.0, results.Single(), 1e-9);
        }

        [MUT.TestMethod]
        public void SetVisioShapeCells_CellsIsPositional()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<double>(
                rect("$s", 0, 0) +
                "$c = New-VisioShapeCells; $c.XFormWidth = '4'; Set-VisioShapeCells $c -Shape $s; " +
                "$s.CellsU('Width').ResultIU");
            MUT.Assert.AreEqual(4.0, results.Single(), 1e-9);
        }

        [MUT.TestMethod]
        public void SetVisioShapeCells_OneCellsObjectForTwoShapes_AppliesToBoth()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$a", 0, 0) + rect("$b", 3, 0) +
                "$c = New-VisioShapeCells; $c.XFormWidth = '2'; Set-VisioShapeCells -Cells $c -Shape @($a, $b); " +
                "\"$($a.CellsU('Width').ResultIU)|$($b.CellsU('Width').ResultIU)\"");
            MUT.Assert.AreEqual("2|2", results.Single());
        }

        [MUT.TestMethod]
        public void SetVisioShapeCells_OneCellsObjectPerShape_AppliesEachInOrder()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$a", 0, 0) + rect("$b", 3, 0) +
                "$c = New-VisioShapeCells -Count 2; $c[0].XFormWidth = '2'; $c[1].XFormWidth = '5'; " +
                "Set-VisioShapeCells -Cells $c -Shape @($a, $b); " +
                "\"$($a.CellsU('Width').ResultIU)|$($b.CellsU('Width').ResultIU)\"");
            MUT.Assert.AreEqual("2|5", results.Single());
        }

        [MUT.TestMethod]
        public void SetVisioShapeCells_WithoutShapeParameter_UsesTheSelection()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                rect("$a", 0, 0) + rect("$b", 3, 0) +
                "Select-VisioShape -Shapes $b; $c = New-VisioShapeCells; $c.XFormWidth = '6'; Set-VisioShapeCells -Cells $c; " +
                "\"$($a.CellsU('Width').ResultIU)|$($b.CellsU('Width').ResultIU)\"");
            MUT.Assert.AreEqual("1|6", results.Single());
        }

        [MUT.TestMethod]
        public void SetVisioShapeCells_BlastGuardsAndTestCircular_AreSwitches()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<double>(
                rect("$s", 0, 0) +
                "$c = New-VisioShapeCells; $c.XFormWidth = '3'; Set-VisioShapeCells -Cells $c -BlastGuards -TestCircular -Shape $s; " +
                "$s.CellsU('Width').ResultIU");
            MUT.Assert.AreEqual(3.0, results.Single(), 1e-9);
        }

        [MUT.TestMethod]
        public void SetVisioShapeCells_WithoutCells_ReportsTheMissingMandatoryParameter()
        {
            var ex = RemoverAndCellsCmdletBindingTests.Session.ExpectFailureInNewDocument(
                rect("$s", 0, 0) + "Set-VisioShapeCells -Shape $s");
            MissingMandatory(ex, "Cells");
        }

        [MUT.TestMethod]
        public void SetVisioShapeCells_CellsOfTheWrongType_FailToBind()
        {
            var ex = RemoverAndCellsCmdletBindingTests.Session.ExpectFailureInNewDocument(
                rect("$s", 0, 0) + "Set-VisioShapeCells -Cells (New-VisioPageCells) -Shape $s");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Cannot bind parameter 'Cells'");
        }

        // -- Set-VisioPageCells --------------------------------------------------------

        [MUT.TestMethod]
        public void SetVisioPageCells_Cells_WritesTheFormulasToThePage()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<double>(
                "$p = Get-VisioPage; $c = New-VisioPageCells; $c.PageWidth = '7'; Set-VisioPageCells -Cells $c -Page $p; " +
                "$p.PageSheet.CellsU('PageWidth').ResultIU");
            MUT.Assert.AreEqual(7.0, results.Single(), 1e-9);
        }

        [MUT.TestMethod]
        public void SetVisioPageCells_WithoutPageParameter_UsesTheActivePage()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<double>(
                "$c = New-VisioPageCells; $c.PageHeight = '5'; Set-VisioPageCells -Cells $c; " +
                "(Get-VisioPage -ActivePage).PageSheet.CellsU('PageHeight').ResultIU");
            MUT.Assert.AreEqual(5.0, results.Single(), 1e-9);
        }

        [MUT.TestMethod]
        public void SetVisioPageCells_OneCellsObjectForTwoPages_AppliesToBoth()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                "$p1 = Get-VisioPage; $p2 = New-VisioPage -Name 'two'; " +
                "$c = New-VisioPageCells; $c.PageWidth = '6'; Set-VisioPageCells -Cells $c -Page @($p1, $p2); " +
                "\"$($p1.PageSheet.CellsU('PageWidth').ResultIU)|$($p2.PageSheet.CellsU('PageWidth').ResultIU)\"");
            MUT.Assert.AreEqual("6|6", results.Single());
        }

        [MUT.TestMethod]
        public void SetVisioPageCells_OneCellsObjectPerPage_AppliesEachInOrder()
        {
            var results = RemoverAndCellsCmdletBindingTests.Session.RunInNewDocument<string>(
                "$p1 = Get-VisioPage; $p2 = New-VisioPage -Name 'two'; " +
                "$c = New-VisioPageCells -Count 2; $c[0].PageWidth = '6'; $c[1].PageWidth = '9'; " +
                "Set-VisioPageCells -Cells $c -Page @($p1, $p2); " +
                "\"$($p1.PageSheet.CellsU('PageWidth').ResultIU)|$($p2.PageSheet.CellsU('PageWidth').ResultIU)\"");
            MUT.Assert.AreEqual("6|9", results.Single());
        }

        [MUT.TestMethod]
        public void SetVisioPageCells_WithoutCells_ReportsTheMissingMandatoryParameter()
        {
            var ex = RemoverAndCellsCmdletBindingTests.Session.ExpectFailureInNewDocument("Set-VisioPageCells");
            MissingMandatory(ex, "Cells");
        }

        [MUT.TestMethod]
        public void SetVisioPageCells_CellsOfTheWrongType_FailToBind()
        {
            var ex = RemoverAndCellsCmdletBindingTests.Session.ExpectFailureInNewDocument(
                "Set-VisioPageCells -Cells (New-VisioShapeCells)");
            MUT.StringAssert.Contains(CmdletScriptExtensions.MessageOf(ex), "Cannot bind parameter 'Cells'");
        }
    }
}
