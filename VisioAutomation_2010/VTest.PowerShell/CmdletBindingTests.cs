using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioPowerShell.Commands.VisioApplication;
using VTest.PowerShell.Framework;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VTest.PowerShell
{
    // Cmdlet parameter-binding regression tests: exercises PowerShell's binder
    // via the runspace path (InvokeScript) rather than direct cmdlet.Invoke,
    // because direct Invoke bypasses the binder.
    //
    // First slice of #173, focused on the regression class that shipped 4.6.1's
    // four bugs (Lock/Unlock-VisioShape switches that didn't bind, Export-VisioShape
    // inverted file-existence check, New-VisioShape polyline / Bezier minimum-point
    // validation that didn't actually throw).

    [MUT.TestClass]
    public class CmdletBindingTests
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
            try { CmdletBindingTests.Session.Cmd_Close_VisioApplication(true); }
            catch (System.Exception) { }
            CmdletBindingTests.Session.CleanUp();
        }

        // -- Lock-VisioShape / Unlock-VisioShape: switch parameters ---------------

        [MUT.TestMethod]
        public void LockVisioShape_MoveXSwitch_BindsAndSetsLockMoveXFormula()
        {
            // Regression: the -MoveX switch (and its peers) used to fail to bind,
            // so Lock-VisioShape -MoveX silently left LockMoveX = "0".
            var doc = CmdletBindingTests.Session.Cmd_New_VisioDocument();
            var shape = CmdletBindingTests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 0.0),
                new VisioAutomation.Core.Point(2.0, 2.0)
            });

            CmdletBindingTests.Session.InvokeScript<object>(
                "Lock-VisioShape -MoveX -Shape $s",
                ("s", new[] { (IVisio.Shape)shape }));

            string movex = ((IVisio.Shape)shape).CellsU["LockMoveX"].FormulaU;
            MUT.Assert.AreEqual("1", movex, "Lock-VisioShape -MoveX should set LockMoveX formula to '1'");

            CmdletBindingTests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        public void UnlockVisioShape_MoveXSwitch_BindsAndSetsLockMoveXFormulaToZero()
        {
            // Regression: the symmetric -MoveX switch on Unlock-VisioShape used to
            // fail to bind, so a previously-locked shape stayed locked.
            var doc = CmdletBindingTests.Session.Cmd_New_VisioDocument();
            var shape = CmdletBindingTests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 0.0),
                new VisioAutomation.Core.Point(2.0, 2.0)
            });

            // First lock, then unlock, to verify the unlock path actually flips back.
            CmdletBindingTests.Session.InvokeScript<object>(
                "Lock-VisioShape -MoveX -Shape $s",
                ("s", new[] { (IVisio.Shape)shape }));
            string after_lock = ((IVisio.Shape)shape).CellsU["LockMoveX"].FormulaU;
            MUT.Assert.AreEqual("1", after_lock, "Pre-condition: lock should set LockMoveX = 1");

            CmdletBindingTests.Session.InvokeScript<object>(
                "Unlock-VisioShape -MoveX -Shape $s",
                ("s", new[] { (IVisio.Shape)shape }));

            string after_unlock = ((IVisio.Shape)shape).CellsU["LockMoveX"].FormulaU;
            MUT.Assert.AreEqual("0", after_unlock, "Unlock-VisioShape -MoveX should set LockMoveX formula to '0'");

            CmdletBindingTests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        // -- Export-VisioShape: file-existence check + -Overwrite -----------------

        [MUT.TestMethod]
        public void ExportVisioShape_TargetExists_NoOverwrite_Throws()
        {
            // Regression: the file-existence check used to be inverted, so
            // Export-VisioShape would happily clobber an existing file even
            // without -Overwrite (or refuse to write to a nonexistent path).
            var doc = CmdletBindingTests.Session.Cmd_New_VisioDocument();
            var shape = CmdletBindingTests.Session.Cmd_New_VisioShape_rectangle(new[]
            {
                new VisioAutomation.Core.Point(0.0, 0.0),
                new VisioAutomation.Core.Point(2.0, 2.0)
            });

            string export_path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "vtest_export_" + System.Guid.NewGuid().ToString("N") + ".png");
            System.IO.File.WriteAllText(export_path, "pre-existing stub content");

            try
            {
                bool threw = false;
                try
                {
                    // $ErrorActionPreference = 'Stop' forces the cmdlet's thrown exception
                    // to propagate as a runtime error rather than landing on the error stream.
                    CmdletBindingTests.Session.InvokeScript<object>(
                        "$ErrorActionPreference = 'Stop'; Export-VisioShape -Filename $f -Shape $s",
                        ("f", export_path),
                        ("s", new[] { (IVisio.Shape)shape }));
                }
                catch (System.Exception)
                {
                    threw = true;
                }
                MUT.Assert.IsTrue(threw, "Export-VisioShape should throw when target exists and -Overwrite is not set");

                // The pre-existing stub content must still be there (the cmdlet must not have clobbered it).
                MUT.Assert.AreEqual(
                    "pre-existing stub content",
                    System.IO.File.ReadAllText(export_path),
                    "Export-VisioShape must not overwrite the file when -Overwrite is not set");
            }
            finally
            {
                if (System.IO.File.Exists(export_path))
                {
                    System.IO.File.Delete(export_path);
                }
            }

            CmdletBindingTests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
        }

        [MUT.TestMethod]
        [MUT.Ignore("Blocked by #164: cmdlets that call TargetShapes.ResolveToSelection through the runspace path hit "
            + "'CommandTarget: application does not match doc.application' because the runspace's Client and the test-host's "
            + "Client disagree on which app/doc is active. The no-Overwrite path above doesn't bite because the cmdlet "
            + "throws at the file-existence check before reaching ResolveToSelection. Re-enable when #164 lands.")]
        public void ExportVisioShape_TargetExists_WithOverwrite_ReplacesFile()
        {
            // Intentionally empty body. Re-implement when #164 is resolved.
        }

        // -- New-VisioShape: polyline / Bezier minimum-point validation -----------

        [MUT.TestMethod]
        public void NewVisioShape_Polyline_FewerThan2Points_Throws()
        {
            // Regression: New-VisioShape -Polyline with 1 point used to silently
            // fall through (Visio refused to draw with insufficient vertices, but
            // the cmdlet's _check_num_Points() validation was a no-op on the
            // throw path before 4.6.1).
            var doc = CmdletBindingTests.Session.Cmd_New_VisioDocument();

            try
            {
                bool threw = false;
                try
                {
                    // $ErrorActionPreference = 'Stop' makes cmdlet-thrown exceptions propagate
                    // as runtime errors rather than landing silently on the error stream.
                    CmdletBindingTests.Session.InvokeScript<object>(
                        "$ErrorActionPreference = 'Stop'; New-VisioShape -Polyline -Points $p",
                        ("p", new[] { new VisioAutomation.Core.Point(1.0, 1.0) }));
                }
                catch (System.Exception)
                {
                    threw = true;
                }
                MUT.Assert.IsTrue(threw, "New-VisioShape -Polyline should throw when given fewer than 2 points");
            }
            finally
            {
                CmdletBindingTests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
            }
        }

        [MUT.TestMethod]
        public void NewVisioShape_Bezier_FewerThan4Points_Throws()
        {
            // Regression: same shape as the polyline test, but Bezier requires
            // at least 4 points (two endpoints + two control points).
            var doc = CmdletBindingTests.Session.Cmd_New_VisioDocument();

            try
            {
                bool threw = false;
                try
                {
                    // $ErrorActionPreference = 'Stop' makes cmdlet-thrown exceptions propagate
                    // as runtime errors rather than landing silently on the error stream.
                    CmdletBindingTests.Session.InvokeScript<object>(
                        "$ErrorActionPreference = 'Stop'; New-VisioShape -Bezier -Points $p",
                        ("p", new[]
                        {
                            new VisioAutomation.Core.Point(0.0, 0.0),
                            new VisioAutomation.Core.Point(1.0, 1.0)
                        }));
                }
                catch (System.Exception)
                {
                    threw = true;
                }
                MUT.Assert.IsTrue(threw, "New-VisioShape -Bezier should throw when given fewer than 4 points");
            }
            finally
            {
                CmdletBindingTests.Session.Cmd_Close_VisioDocument(VTestPsArray.From(doc), true);
            }
        }
    }
}
