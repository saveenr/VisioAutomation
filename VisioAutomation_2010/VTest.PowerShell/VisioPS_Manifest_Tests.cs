using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using SMA = System.Management.Automation;

namespace VTest.PowerShell
{
    // Drift check between Visio.psd1's CmdletsToExport list and the cmdlets
    // VisioPS.dll actually exports. Mirrors the same check publish-psmodule.yml
    // runs at publish time, but at unit-test time so a missing or stale entry
    // is caught on every test run instead of only at release time.
    //
    // Pure metadata: no Visio runtime, no PS runspace, no [ClassInitialize].

    [MUT.TestClass]
    public class VisioPS_Manifest_Tests
    {
        private static HashSet<string> GetActualCmdletNames()
        {
            var asm = typeof(VisioPowerShell.Commands.VisioCmdlet).Assembly;
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var t in asm.GetTypes())
            {
                if (t.IsAbstract) continue;
                if (!typeof(SMA.Cmdlet).IsAssignableFrom(t)) continue;
                var attr = t.GetCustomAttribute<SMA.CmdletAttribute>();
                if (attr == null) continue;
                names.Add($"{attr.VerbName}-{attr.NounName}");
            }
            return names;
        }

        private static HashSet<string> GetDeclaredCmdletNames()
        {
            // Visio.psd1 is copied next to the test assembly via VTest.PowerShell.csproj.
            var asmDir = Path.GetDirectoryName(typeof(VisioPS_Manifest_Tests).Assembly.Location);
            var psd1Path = Path.Combine(asmDir, "Visio.psd1");
            MUT.Assert.IsTrue(
                File.Exists(psd1Path),
                $"Expected Visio.psd1 next to test assembly at: {psd1Path}. Check the <None Include=\"..\\VisioPowerShell\\Visio.psd1\" Link=\"Visio.psd1\" CopyToOutputDirectory=\"PreserveNewest\" /> entry in VTest.PowerShell.csproj.");

            var content = File.ReadAllText(psd1Path);
            var tokens = SMA.PSParser.Tokenize(content, out var errors);
            MUT.Assert.AreEqual(
                0,
                errors.Count,
                $"Visio.psd1 has {errors.Count} PowerShell parse error(s); fix the manifest before running this test.");

            // Walk tokens to find the CmdletsToExport member, then collect
            // every String token inside its @(...) group. Tracks group depth
            // to handle the (theoretical) case of nested parentheses inside
            // the array.
            for (int i = 0; i < tokens.Count; i++)
            {
                if (tokens[i].Type != SMA.PSTokenType.Member ||
                    !string.Equals(tokens[i].Content, "CmdletsToExport", StringComparison.Ordinal))
                {
                    continue;
                }

                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int depth = 0;
                bool sawOpen = false;
                for (int j = i + 1; j < tokens.Count; j++)
                {
                    var tok = tokens[j];
                    if (tok.Type == SMA.PSTokenType.GroupStart)
                    {
                        depth++;
                        sawOpen = true;
                    }
                    else if (tok.Type == SMA.PSTokenType.GroupEnd)
                    {
                        depth--;
                        if (sawOpen && depth == 0) return names;
                    }
                    else if (tok.Type == SMA.PSTokenType.String && depth > 0)
                    {
                        names.Add(tok.Content);
                    }
                }
                MUT.Assert.Fail("Found CmdletsToExport but could not match its closing ')'.");
            }
            MUT.Assert.Fail("Did not find 'CmdletsToExport' member in Visio.psd1.");
            return null; // unreachable
        }

        [MUT.TestMethod]
        public void VisioPS_Manifest_CmdletsToExport_HasNoMissingEntries()
        {
            var actual = GetActualCmdletNames();
            var declared = GetDeclaredCmdletNames();
            var missing = actual.Except(declared, StringComparer.OrdinalIgnoreCase)
                                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                                .ToArray();
            MUT.Assert.AreEqual(
                0,
                missing.Length,
                $"Cmdlet(s) defined in VisioPS.dll but missing from Visio.psd1's CmdletsToExport: {string.Join(", ", missing)}. Add them to the manifest list in VisioAutomation_2010/VisioPowerShell/Visio.psd1.");
        }

        [MUT.TestMethod]
        public void VisioPS_Manifest_CmdletsToExport_HasNoExtraEntries()
        {
            var actual = GetActualCmdletNames();
            var declared = GetDeclaredCmdletNames();
            var extra = declared.Except(actual, StringComparer.OrdinalIgnoreCase)
                                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                                .ToArray();
            MUT.Assert.AreEqual(
                0,
                extra.Length,
                $"Visio.psd1's CmdletsToExport lists name(s) that are not actual cmdlets in VisioPS.dll: {string.Join(", ", extra)}. Remove them from the manifest list in VisioAutomation_2010/VisioPowerShell/Visio.psd1, or restore the cmdlet class.");
        }
    }
}
