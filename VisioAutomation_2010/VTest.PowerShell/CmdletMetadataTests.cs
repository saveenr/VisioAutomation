using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using SMA = System.Management.Automation;

namespace VTest.PowerShell
{
    // Breadth check for #263: invariants every cmdlet in VisioPS.dll must satisfy, checked by reflection
    // over all of them at once. It catches declaration mistakes (a reused position number, a misspelled
    // default parameter set, a verb PowerShell does not recognize) in cmdlets that have no behavior test
    // yet. It says nothing about what a cmdlet does; the per-cmdlet binding tests do that.
    //
    // Pure metadata: no Visio runtime, no runspace, no [ClassInitialize].
    [MUT.TestClass]
    public class CmdletMetadataTests
    {
        private static IEnumerable<Type> CmdletTypes()
        {
            var asm = typeof(VisioPowerShell.Commands.VisioCmdlet).Assembly;
            return asm.GetTypes()
                .Where(t => !t.IsAbstract && typeof(SMA.Cmdlet).IsAssignableFrom(t) && t.GetCustomAttribute<SMA.CmdletAttribute>() != null)
                .OrderBy(t => t.Name, StringComparer.Ordinal)
                .ToList();
        }

        private sealed class Param
        {
            public string Name;
            public SMA.ParameterAttribute[] Attributes;
        }

        // Public instance fields and properties that carry at least one [Parameter].
        private static List<Param> ParametersOf(Type t)
        {
            var list = new List<Param>();
            foreach (var member in t.GetMembers(BindingFlags.Public | BindingFlags.Instance).Where(m => m is FieldInfo || m is PropertyInfo))
            {
                var attrs = member.GetCustomAttributes<SMA.ParameterAttribute>(true).ToArray();
                if (attrs.Length > 0)
                {
                    list.Add(new Param { Name = member.Name, Attributes = attrs });
                }
            }
            return list;
        }

        private static string Describe(Type t)
        {
            var c = t.GetCustomAttribute<SMA.CmdletAttribute>();
            return c.VerbName + "-" + c.NounName;
        }

        private static HashSet<string> ApprovedVerbs()
        {
            // PowerShell's own verb constant classes are the list Get-Verb reports.
            var verbs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var asm = typeof(SMA.VerbsCommon).Assembly;
            foreach (var t in asm.GetTypes().Where(x => x.IsPublic && x.Namespace == "System.Management.Automation" && x.Name.StartsWith("Verbs", StringComparison.Ordinal)))
            {
                foreach (var f in t.GetFields(BindingFlags.Public | BindingFlags.Static).Where(f => f.FieldType == typeof(string)))
                {
                    verbs.Add((string)f.GetValue(null));
                }
            }
            return verbs;
        }

        private static void AssertNoProblems(string what, List<string> problems)
        {
            MUT.Assert.AreEqual(0, problems.Count, what + ":\n  " + string.Join("\n  ", problems));
        }

        [MUT.TestMethod]
        public void Cmdlets_AreFoundByReflection()
        {
            // Guards the other tests in this class: they pass trivially if the scan finds nothing.
            MUT.Assert.IsTrue(CmdletTypes().Count() >= 60, "found only " + CmdletTypes().Count() + " cmdlet types");
        }

        [MUT.TestMethod]
        public void Cmdlets_AllDeriveFromVisioCmdlet()
        {
            var problems = CmdletTypes()
                .Where(t => !typeof(VisioPowerShell.Commands.VisioCmdlet).IsAssignableFrom(t))
                .Select(Describe).ToList();
            AssertNoProblems("cmdlets that do not derive from VisioCmdlet (so they miss the shared client and error handling)", problems);
        }

        [MUT.TestMethod]
        public void Cmdlets_AllUseAnApprovedVerb()
        {
            var approved = ApprovedVerbs();
            MUT.Assert.IsTrue(approved.Count > 50, "verb list looks wrong: " + approved.Count);
            var problems = CmdletTypes()
                .Where(t => !approved.Contains(t.GetCustomAttribute<SMA.CmdletAttribute>().VerbName))
                .Select(Describe).ToList();
            AssertNoProblems("cmdlets whose verb is not an approved PowerShell verb (Import-Module warns about these)", problems);
        }

        [MUT.TestMethod]
        public void Cmdlets_NamesAreUnique()
        {
            var problems = CmdletTypes()
                .GroupBy(t => Describe(t), StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key + " is declared by " + string.Join(", ", g.Select(t => t.Name))).ToList();
            AssertNoProblems("duplicate cmdlet names", problems);
        }

        [MUT.TestMethod]
        public void Cmdlets_ParameterNamesAreUniqueIgnoringCase()
        {
            var problems = new List<string>();
            foreach (var t in CmdletTypes())
            {
                foreach (var g in ParametersOf(t).GroupBy(p => p.Name, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                {
                    problems.Add(Describe(t) + ": " + string.Join(" / ", g.Select(p => p.Name)));
                }
            }
            AssertNoProblems("parameters whose names differ only by case (PowerShell cannot tell them apart)", problems);
        }

        // Cmdlets with a known position clash, each tied to the issue that will fix it. Remove an entry
        // when its issue is fixed: Cmdlets_KnownPositionClashes_StillExist fails once an entry is stale.
        private static readonly Dictionary<string, string> KnownPositionClashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // (none at the moment: Format-VisioWindow, #290, was fixed.)
        };

        private static List<string> PositionClashes(Type t)
        {
            var problems = new List<string>();
            {
                var parameters = ParametersOf(t);
                // A [Parameter] without a set name belongs to every set, so it competes with each named set.
                var sets = parameters.SelectMany(p => p.Attributes).Select(a => a.ParameterSetName)
                    .Where(n => n != SMA.ParameterAttribute.AllParameterSets).Distinct().ToList();
                if (sets.Count == 0) { sets.Add(SMA.ParameterAttribute.AllParameterSets); }

                foreach (var set in sets)
                {
                    var positions = new List<KeyValuePair<int, string>>();
                    foreach (var p in parameters)
                    {
                        foreach (var a in p.Attributes.Where(a => a.Position != int.MinValue &&
                            (a.ParameterSetName == set || a.ParameterSetName == SMA.ParameterAttribute.AllParameterSets)))
                        {
                            positions.Add(new KeyValuePair<int, string>(a.Position, p.Name));
                        }
                    }
                    foreach (var dup in positions.GroupBy(x => x.Key).Where(g => g.Count() > 1))
                    {
                        problems.Add(Describe(t) + " (set '" + set + "'): position " + dup.Key + " is used by " + string.Join(", ", dup.Select(x => x.Value)));
                    }
                }
            }
            return problems;
        }

        [MUT.TestMethod]
        public void Cmdlets_NoParameterSetReusesAPositionNumber()
        {
            var problems = CmdletTypes()
                .Where(t => !KnownPositionClashes.ContainsKey(Describe(t)))
                .SelectMany(PositionClashes).ToList();
            AssertNoProblems("positions used by more than one parameter in the same parameter set (a clash makes the later parameter unreachable by position)", problems);
        }

        [MUT.TestMethod]
        public void Cmdlets_KnownPositionClashes_StillExist()
        {
            var types = CmdletTypes().ToDictionary(Describe, StringComparer.OrdinalIgnoreCase);
            var stale = new List<string>();
            foreach (var known in KnownPositionClashes)
            {
                if (!types.ContainsKey(known.Key) || PositionClashes(types[known.Key]).Count == 0)
                {
                    stale.Add(known.Key + " (" + known.Value + ")");
                }
            }
            AssertNoProblems("known position clashes that no longer exist: remove them from KnownPositionClashes", stale);
        }

        [MUT.TestMethod]
        public void Cmdlets_DefaultParameterSetNameMatchesADeclaredSet()
        {
            var problems = new List<string>();
            foreach (var t in CmdletTypes())
            {
                var def = t.GetCustomAttribute<SMA.CmdletAttribute>().DefaultParameterSetName;
                if (string.IsNullOrEmpty(def)) { continue; }
                var declared = ParametersOf(t).SelectMany(p => p.Attributes).Select(a => a.ParameterSetName);
                if (!declared.Contains(def, StringComparer.Ordinal))
                {
                    problems.Add(Describe(t) + ": default set '" + def + "' is not used by any parameter");
                }
            }
            AssertNoProblems("default parameter sets that no parameter belongs to (usually a typo)", problems);
        }

        [MUT.TestMethod]
        public void Cmdlets_WithSeveralParameterSets_DeclareADefaultOrOnlyUseDistinguishingParameters()
        {
            // If a cmdlet has more than one named set and no default, a call that supplies none of the
            // set-specific parameters cannot be resolved. This lists them so each is a deliberate choice;
            // the set below is the reviewed list. A new cmdlet that appears here needs a default set, or
            // a mandatory parameter in every set, and then an entry here.
            var reviewed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var problems = new List<string>();
            foreach (var t in CmdletTypes())
            {
                var named = ParametersOf(t).SelectMany(p => p.Attributes).Select(a => a.ParameterSetName)
                    .Where(n => n != SMA.ParameterAttribute.AllParameterSets).Distinct().ToList();
                if (named.Count < 2) { continue; }
                if (!string.IsNullOrEmpty(t.GetCustomAttribute<SMA.CmdletAttribute>().DefaultParameterSetName)) { continue; }

                // Fine when every set has at least one mandatory parameter of its own.
                var everySetHasAMandatory = named.All(set => ParametersOf(t)
                    .SelectMany(p => p.Attributes)
                    .Any(a => a.ParameterSetName == set && a.Mandatory));
                if (everySetHasAMandatory) { continue; }

                if (!reviewed.Contains(Describe(t)))
                {
                    problems.Add(Describe(t) + " has sets [" + string.Join(", ", named) + "], no default set, and a set with no mandatory parameter");
                }
            }
            AssertNoProblems("cmdlets with several parameter sets that may be ambiguous", problems);
        }
    }
}
