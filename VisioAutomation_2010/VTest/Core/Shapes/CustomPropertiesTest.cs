using System.Globalization;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes;
using VA=VisioAutomation;

namespace VTest.Core.Shapes
{
    [MUT.TestClass]
    public class CustomPropertiesTest : Framework.VTest
    {
        [MUT.TestMethod]
        public void CustomProps_Names()
        {
            MUT.Assert.IsFalse(VA.Shapes.CustomPropertyHelper.IsValidName(null));
            MUT.Assert.IsFalse(VA.Shapes.CustomPropertyHelper.IsValidName(string.Empty));
            MUT.Assert.IsFalse(VA.Shapes.CustomPropertyHelper.IsValidName(" foo "));
            MUT.Assert.IsFalse(VA.Shapes.CustomPropertyHelper.IsValidName("foo "));
            MUT.Assert.IsFalse(VA.Shapes.CustomPropertyHelper.IsValidName("foo\t"));
            MUT.Assert.IsFalse(VA.Shapes.CustomPropertyHelper.IsValidName("fo bar"));
            MUT.Assert.IsTrue(VA.Shapes.CustomPropertyHelper.IsValidName("foobar"));
        }

        [MUT.TestMethod]
        public void SimpleCP()
        {
            var page1 = this.GetNewPage();

            // Draw a shape
            var s1 = page1.DrawRectangle(1, 1, 4, 3);

            int cp_type = 0; // string type

            // Set some properties on it
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO1", "\"BAR1\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO2", "\"BAR2\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO3", "\"BAR3\"", cp_type);

            // Delete one of those properties
            VA.Shapes.CustomPropertyHelper.Delete(s1, "FOO2");

            // Set the value of an existing properties
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO3", "\"BAR3updated\"", cp_type);

            // retrieve all the properties
            var props = VA.Shapes.CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);

            var cp_foo1 = props["FOO1"];
            // var cp_foo2 = props["FOO2"]; there is no prop called FOO2
            var cp_foo3 = props["FOO3"];

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }

        [MUT.TestMethod]
        public void CustomProps_CRUD()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            s1.Text = "Checking for Custom Properties";

            // A new rectangle should have zero props
            var c0 = VA.Shapes.CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(0, c0.Count);


            int cp_type = 0; // 0 for string

            // Set one property
            // Notice that the properties some back double-quoted
            VA.Shapes.CustomPropertyHelper.Set(s1, "PROP1", "\"VAL1\"", cp_type);

            var c1 = VA.Shapes.CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);

            MUT.Assert.AreEqual(1, c1.Count);
            MUT.Assert.IsTrue(c1.ContainsKey("PROP1"));
            MUT.Assert.AreEqual("\"VAL1\"", c1["PROP1"].Value.Value);

            // Add another property
            VA.Shapes.CustomPropertyHelper.Set(s1, "PROP2", "\"VAL 2\"", cp_type);
            var c2 = VA.Shapes.CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);

            MUT.Assert.AreEqual(2, c2.Count);
            MUT.Assert.IsTrue(c2.ContainsKey("PROP1"));
            MUT.Assert.AreEqual("\"VAL1\"", c2["PROP1"].Value.Value);
            MUT.Assert.IsTrue(c2.ContainsKey("PROP2"));
            MUT.Assert.AreEqual("\"VAL 2\"", c2["PROP2"].Value.Value);

            // Modify the value of the second property
            VA.Shapes.CustomPropertyHelper.Set(s1, "PROP2", "\"VAL 2 MOD\"", cp_type);
            var c3 = VA.Shapes.CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);
  
            MUT.Assert.AreEqual(2, c3.Count);
            MUT.Assert.IsTrue(c3.ContainsKey("PROP1"));
            MUT.Assert.AreEqual("\"VAL1\"", c3["PROP1"].Value.Value);
            MUT.Assert.IsTrue(c3.ContainsKey("PROP2"));
            MUT.Assert.AreEqual("\"VAL 2 MOD\"", c3["PROP2"].Value.Value);

            // Now delete all the custom properties
            foreach (string name in c3.Keys)
            {
                VA.Shapes.CustomPropertyHelper.Delete(s1, name);
            }

            var c4 = VA.Shapes.CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);


            MUT.Assert.AreEqual(0, c4.Count);

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }

        [MUT.TestMethod]
        public void CustomProps_AllTypes()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            s1.Text = "Checking for Custom Properties";
            
            // String Custom Property
            var prop_string_in = new CustomPropertyCells();
            prop_string_in.Format = "\"Format\"";
            prop_string_in.Label = "\"Label\"";
            prop_string_in.Prompt = "\"Prompt\"";
            prop_string_in.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.String);
            prop_string_in.Value = "1";

            // Boolean
            var prop_bool_in = new CustomPropertyCells();
            prop_bool_in.Format = "\"Format\"";
            prop_bool_in.Label = "\"Label\"";
            prop_bool_in.Prompt = "\"Prompt\"";
            prop_bool_in.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Boolean);
            prop_bool_in.Value = true;

            // Date
            var dt = new System.DateTime(2017,3,31,14,5,6);
            var st = dt.ToString(CultureInfo.InvariantCulture);
            var prop_date_in = new CustomPropertyCells();
            prop_date_in.Format = "\"Format\"";
            prop_date_in.Label = "\"Label\"";
            prop_date_in.Prompt = "\"Prompt\"";
            prop_date_in.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Date);
            prop_date_in.Value = string.Format("DATETIME(\"{0}\")", st); ;

            // Boolean
            var prop_number_in = new CustomPropertyCells();
            prop_number_in.Format = "\"Format\"";
            prop_number_in.Label = "\"Label\"";
            prop_number_in.Prompt = "\"Prompt\"";
            prop_number_in.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Number);
            prop_number_in.Value = "3.14";

            VA.Shapes.CustomPropertyHelper.Set(s1, "PROP_STRING", prop_string_in);
            VA.Shapes.CustomPropertyHelper.Set(s1, "PROP_BOOLEAN", prop_bool_in);
            VA.Shapes.CustomPropertyHelper.Set(s1, "PROP_DATE", prop_date_in);
            VA.Shapes.CustomPropertyHelper.Set(s1, "PROP_NUMBER", prop_number_in);

            var props_dic = VA.Shapes.CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);


            var prop_string_out = props_dic["PROP_STRING"];

            MUT.Assert.AreEqual("\"Format\"", prop_string_out.Format.Value);
            MUT.Assert.AreEqual("\"Label\"", prop_string_out.Label.Value);
            MUT.Assert.AreEqual("\"Prompt\"", prop_string_out.Prompt.Value);
            MUT.Assert.AreEqual("0", prop_string_out.Type.Value);
            MUT.Assert.AreEqual("1", prop_string_out.Value.Value);

            var prop_bool_out = props_dic["PROP_BOOLEAN"];
            MUT.Assert.AreEqual("\"Format\"", prop_bool_out.Format.Value);
            MUT.Assert.AreEqual("\"Label\"", prop_bool_out.Label.Value);
            MUT.Assert.AreEqual("\"Prompt\"", prop_bool_out.Prompt.Value);
            MUT.Assert.AreEqual("3", prop_bool_out.Type.Value);
            MUT.Assert.AreEqual("TRUE", prop_bool_out.Value.Value);

            var prop_date_out = props_dic["PROP_DATE"];
            MUT.Assert.AreEqual("\"Format\"", prop_date_out.Format.Value);
            MUT.Assert.AreEqual("\"Label\"", prop_date_out.Label.Value);
            MUT.Assert.AreEqual("\"Prompt\"", prop_date_out.Prompt.Value);
            MUT.Assert.AreEqual("5", prop_date_out.Type.Value);
            MUT.Assert.AreEqual("DATETIME(\"03/31/2017 14:05:06\")", prop_date_out.Value.Value);

            var prop_number_out = props_dic["PROP_NUMBER"];
            MUT.Assert.AreEqual("\"Format\"", prop_number_out.Format.Value);
            MUT.Assert.AreEqual("\"Label\"", prop_number_out.Label.Value);
            MUT.Assert.AreEqual("\"Prompt\"", prop_number_out.Prompt.Value);
            MUT.Assert.AreEqual("2", prop_number_out.Type.Value);
            MUT.Assert.AreEqual("3.14", prop_number_out.Value.Value);

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }

        // Issue #144 characterization tests — what does Visio actually do when
        // CustomPropertyCells fields are set to un-encoded inputs and Set is called?
        // Each [Type=X]Characterization test below locks in current behavior so any
        // future fix change surfaces as a deliberate test failure.
        // Full discussion + cross-Type matrix: docs/internal/custom-property-encoding.md.

        [MUT.TestMethod]
        public void CustomProps_UnencodedValueCharacterization()
        {
            // Type=String. Behavior matrix (current Visio install, locked 2026-05-06):
            //
            // Input                                  | Outcome
            // ---------------------------------------+---------------------------------------
            // "testVal" plain identifier             | THROWS COMException #NAME?
            // "42"  numeric-looking                  | succeeds, formula=42, result=42.0000
            // "hello world" spaces                   | THROWS COMException #NAME?
            // ""    empty unquoted                   | succeeds, formula=[empty], result=0.0000
            // "\"\""  empty quoted                   | round-trips, formula=\"\", result=[empty]
            // null                                   | HasValue=false, cell unwritten; default formula=0, result=0.0000
            // "\"testVal\"" pre-quoted               | round-trips, formula=\"testVal\", result=testVal
            // " "   single space unquoted            | succeeds, formula=[empty], result=0.0000
            // "\" \"" single space quoted            | round-trips, formula=\" \", result=[space]
            //
            // Unencoded Label / Format / Prompt with a plain identifier value also
            // throw COMException #NAME? regardless of Type; the string-typed
            // constructors propagate the same trap to .Value.

            var page1 = this.GetNewPage();
            var failures = new System.Collections.Generic.List<string>();
            int caseIndex = 0;

            string MakePropName() { caseIndex++; return "P" + caseIndex; }

            void RunOK(string label, System.Action<CustomPropertyCells> setup, string expFormula, string expResult)
            {
                string propName = MakePropName();
                var s = page1.DrawRectangle(0, 0, 1, 1);
                var cp = new CustomPropertyCells();
                cp.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.String);
                try
                {
                    setup(cp);
                    VA.Shapes.CustomPropertyHelper.Set(s, propName, cp);
                    var fdic = VA.Shapes.CustomPropertyHelper.GetDictionary(s, VisioAutomation.Core.CellValueType.Formula);
                    var rdic = VA.Shapes.CustomPropertyHelper.GetDictionary(s, VisioAutomation.Core.CellValueType.Result);
                    string af = fdic.ContainsKey(propName) ? (fdic[propName].Value.Value ?? "<null>") : "<missing>";
                    string ar = rdic.ContainsKey(propName) ? (rdic[propName].Value.Value ?? "<null>") : "<missing>";
                    if (af != expFormula || ar != expResult)
                    {
                        failures.Add(string.Format("[{0}] exp formula=[{1}] result=[{2}], got formula=[{3}] result=[{4}]",
                            label, expFormula, expResult, af, ar));
                    }
                }
                catch (System.Exception ex)
                {
                    failures.Add(string.Format("[{0}] expected success but THREW {1}: {2}", label, ex.GetType().Name, ex.Message));
                }
            }

            void RunThrows(string label, System.Action<CustomPropertyCells> setup, string expMsgContains)
            {
                string propName = MakePropName();
                var s = page1.DrawRectangle(0, 0, 1, 1);
                var cp = new CustomPropertyCells();
                cp.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.String);
                try
                {
                    setup(cp);
                    VA.Shapes.CustomPropertyHelper.Set(s, propName, cp);
                    failures.Add(string.Format("[{0}] expected COMException with [{1}] but Set succeeded", label, expMsgContains));
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    if (!ex.Message.Contains(expMsgContains))
                    {
                        failures.Add(string.Format("[{0}] threw COMException but message [{1}] doesn't contain [{2}]", label, ex.Message, expMsgContains));
                    }
                }
                catch (System.Exception ex)
                {
                    failures.Add(string.Format("[{0}] expected COMException but threw {1}: {2}", label, ex.GetType().Name, ex.Message));
                }
            }

            void RunCtorThrows(string label, System.Func<CustomPropertyCells> ctor, string expMsgContains)
            {
                string propName = MakePropName();
                var s = page1.DrawRectangle(0, 0, 1, 1);
                try
                {
                    var cp = ctor();
                    VA.Shapes.CustomPropertyHelper.Set(s, propName, cp);
                    failures.Add(string.Format("[{0}] expected COMException with [{1}] but Set succeeded", label, expMsgContains));
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    if (!ex.Message.Contains(expMsgContains))
                    {
                        failures.Add(string.Format("[{0}] threw COMException but message [{1}] doesn't contain [{2}]", label, ex.Message, expMsgContains));
                    }
                }
                catch (System.Exception ex)
                {
                    failures.Add(string.Format("[{0}] expected COMException but threw {1}: {2}", label, ex.GetType().Name, ex.Message));
                }
            }

            // === Type=String, Value field ===
            RunThrows("C1 plain identifier", cp => cp.Value = "testVal", "#NAME?");
            RunOK("C2 numeric string", cp => cp.Value = "42", "42", "42.0000");
            RunThrows("C3 spaces in middle", cp => cp.Value = "hello world", "#NAME?");
            RunOK("C4a empty unquoted", cp => cp.Value = "", "", "0.0000");
            RunOK("C4b empty quoted", cp => cp.Value = "\"\"", "\"\"", "");
            RunOK("C5 null Value (cell unwritten, Visio default)", cp => cp.Value = (string)null, "0", "0.0000");
            RunOK("C6 pre-quoted plain", cp => cp.Value = "\"testVal\"", "\"testVal\"", "testVal");
            RunOK("C7a single space unquoted", cp => cp.Value = " ", "", "0.0000");
            RunOK("C7b single space quoted", cp => cp.Value = "\" \"", "\" \"", " ");

            // === Other string-formula fields with an unencoded plain identifier ===
            RunThrows("L1 unencoded Label", cp => { cp.Value = "\"v\""; cp.Label = "labelVal"; }, "#NAME?");
            RunThrows("F1 unencoded Format", cp => { cp.Value = "\"v\""; cp.Format = "formatVal"; }, "#NAME?");
            RunThrows("P1 unencoded Prompt", cp => { cp.Value = "\"v\""; cp.Prompt = "promptVal"; }, "#NAME?");

            // === String-typed constructors propagate the trap to .Value ===
            RunCtorThrows("K1 ctor(string)", () => new CustomPropertyCells("testVal"), "#NAME?");
            RunCtorThrows("K2 ctor(string, CustomPropertyType.String)", () => new CustomPropertyCells("testVal", CustomPropertyType.String), "#NAME?");

            if (failures.Count > 0)
            {
                MUT.Assert.Fail("\n" + string.Join("\n", failures));
            }

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }

        [MUT.TestMethod]
        public void CustomProps_NumberTypeCharacterization()
        {
            // Type=Number (Type=2). Behavior matrix (locked 2026-05-06):
            //
            // Input             | Outcome
            // ------------------+-------------------------------------------
            // "42"              | succeeds, formula=42, result=42.0000
            // "3.14"            | succeeds, formula=3.14, result=3.1400
            // "testVal"         | THROWS COMException #NAME?
            // "\"42\""          | succeeds, formula=\"42\", result=42 (quoted accepted, unquoted in Result)
            // ""    empty       | succeeds, formula=[empty], result=0.0000
            // null              | HasValue=false, cell unwritten; default formula=0, result=0.0000

            var page1 = this.GetNewPage();
            var failures = new System.Collections.Generic.List<string>();
            int caseIndex = 0;

            string MakePropName() { caseIndex++; return "N" + caseIndex; }

            void RunOK(string label, System.Action<CustomPropertyCells> setup, string expFormula, string expResult)
            {
                string propName = MakePropName();
                var s = page1.DrawRectangle(0, 0, 1, 1);
                var cp = new CustomPropertyCells();
                cp.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Number);
                try
                {
                    setup(cp);
                    VA.Shapes.CustomPropertyHelper.Set(s, propName, cp);
                    var fdic = VA.Shapes.CustomPropertyHelper.GetDictionary(s, VisioAutomation.Core.CellValueType.Formula);
                    var rdic = VA.Shapes.CustomPropertyHelper.GetDictionary(s, VisioAutomation.Core.CellValueType.Result);
                    string af = fdic.ContainsKey(propName) ? (fdic[propName].Value.Value ?? "<null>") : "<missing>";
                    string ar = rdic.ContainsKey(propName) ? (rdic[propName].Value.Value ?? "<null>") : "<missing>";
                    if (af != expFormula || ar != expResult)
                    {
                        failures.Add(string.Format("[{0}] exp formula=[{1}] result=[{2}], got formula=[{3}] result=[{4}]",
                            label, expFormula, expResult, af, ar));
                    }
                }
                catch (System.Exception ex)
                {
                    failures.Add(string.Format("[{0}] expected success but THREW {1}: {2}", label, ex.GetType().Name, ex.Message));
                }
            }

            void RunThrows(string label, System.Action<CustomPropertyCells> setup, string expMsgContains)
            {
                string propName = MakePropName();
                var s = page1.DrawRectangle(0, 0, 1, 1);
                var cp = new CustomPropertyCells();
                cp.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Number);
                try
                {
                    setup(cp);
                    VA.Shapes.CustomPropertyHelper.Set(s, propName, cp);
                    failures.Add(string.Format("[{0}] expected COMException with [{1}] but Set succeeded", label, expMsgContains));
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    if (!ex.Message.Contains(expMsgContains))
                    {
                        failures.Add(string.Format("[{0}] threw COMException but message [{1}] doesn't contain [{2}]", label, ex.Message, expMsgContains));
                    }
                }
                catch (System.Exception ex)
                {
                    failures.Add(string.Format("[{0}] expected COMException but threw {1}: {2}", label, ex.GetType().Name, ex.Message));
                }
            }

            RunOK("N1 numeric integer string", cp => cp.Value = "42", "42", "42.0000");
            RunOK("N2 numeric decimal string", cp => cp.Value = "3.14", "3.14", "3.1400");
            RunThrows("N3 plain identifier", cp => cp.Value = "testVal", "#NAME?");
            RunOK("N4 quoted numeric", cp => cp.Value = "\"42\"", "\"42\"", "42");
            RunOK("N5 empty unquoted", cp => cp.Value = "", "", "0.0000");
            RunOK("N6 null Value (cell unwritten, Visio default)", cp => cp.Value = (string)null, "0", "0.0000");

            if (failures.Count > 0)
            {
                MUT.Assert.Fail("\n" + string.Join("\n", failures));
            }

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }

        [MUT.TestMethod]
        public void CustomProps_BooleanTypeCharacterization()
        {
            // Type=Boolean (Type=3). Behavior matrix (locked 2026-05-06):
            //
            // Input          | Outcome
            // ---------------+-------------------------------------------
            // true (literal) | succeeds, formula=TRUE, result=TRUE
            // false (literal)| succeeds, formula=FALSE, result=FALSE
            // "TRUE"         | succeeds, formula=TRUE, result=TRUE
            // "FALSE"        | succeeds, formula=FALSE, result=FALSE
            // "true" lower   | succeeds, normalised to formula=TRUE, result=TRUE
            // "1"            | succeeds, formula=1, result=1.0000  (NUMERIC, not bool — Type metadata says Boolean but Result is number)
            // "0"            | succeeds, formula=0, result=0.0000  (same — numeric Result despite Type=Boolean)
            // "BAR" plain id | THROWS COMException #NAME?
            // ""    empty    | succeeds, formula=[empty], result=0.0000
            // null           | HasValue=false, cell unwritten; default formula=0, result=0.0000

            var page1 = this.GetNewPage();
            var failures = new System.Collections.Generic.List<string>();
            int caseIndex = 0;

            string MakePropName() { caseIndex++; return "B" + caseIndex; }

            void RunOK(string label, System.Action<CustomPropertyCells> setup, string expFormula, string expResult)
            {
                string propName = MakePropName();
                var s = page1.DrawRectangle(0, 0, 1, 1);
                var cp = new CustomPropertyCells();
                cp.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Boolean);
                try
                {
                    setup(cp);
                    VA.Shapes.CustomPropertyHelper.Set(s, propName, cp);
                    var fdic = VA.Shapes.CustomPropertyHelper.GetDictionary(s, VisioAutomation.Core.CellValueType.Formula);
                    var rdic = VA.Shapes.CustomPropertyHelper.GetDictionary(s, VisioAutomation.Core.CellValueType.Result);
                    string af = fdic.ContainsKey(propName) ? (fdic[propName].Value.Value ?? "<null>") : "<missing>";
                    string ar = rdic.ContainsKey(propName) ? (rdic[propName].Value.Value ?? "<null>") : "<missing>";
                    if (af != expFormula || ar != expResult)
                    {
                        failures.Add(string.Format("[{0}] exp formula=[{1}] result=[{2}], got formula=[{3}] result=[{4}]",
                            label, expFormula, expResult, af, ar));
                    }
                }
                catch (System.Exception ex)
                {
                    failures.Add(string.Format("[{0}] expected success but THREW {1}: {2}", label, ex.GetType().Name, ex.Message));
                }
            }

            void RunThrows(string label, System.Action<CustomPropertyCells> setup, string expMsgContains)
            {
                string propName = MakePropName();
                var s = page1.DrawRectangle(0, 0, 1, 1);
                var cp = new CustomPropertyCells();
                cp.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Boolean);
                try
                {
                    setup(cp);
                    VA.Shapes.CustomPropertyHelper.Set(s, propName, cp);
                    failures.Add(string.Format("[{0}] expected COMException with [{1}] but Set succeeded", label, expMsgContains));
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    if (!ex.Message.Contains(expMsgContains))
                    {
                        failures.Add(string.Format("[{0}] threw COMException but message [{1}] doesn't contain [{2}]", label, ex.Message, expMsgContains));
                    }
                }
                catch (System.Exception ex)
                {
                    failures.Add(string.Format("[{0}] expected COMException but threw {1}: {2}", label, ex.GetType().Name, ex.Message));
                }
            }

            RunOK("B1 literal bool true", cp => cp.Value = true, "TRUE", "TRUE");
            RunOK("B2 literal bool false", cp => cp.Value = false, "FALSE", "FALSE");
            RunOK("B3 string TRUE upper", cp => cp.Value = "TRUE", "TRUE", "TRUE");
            RunOK("B4 string FALSE upper", cp => cp.Value = "FALSE", "FALSE", "FALSE");
            RunOK("B5 string true lower normalises to TRUE", cp => cp.Value = "true", "TRUE", "TRUE");
            RunOK("B6 string 1 (numeric, Type metadata mismatch)", cp => cp.Value = "1", "1", "1.0000");
            RunOK("B7 string 0 (numeric, Type metadata mismatch)", cp => cp.Value = "0", "0", "0.0000");
            RunThrows("B8 plain identifier", cp => cp.Value = "BAR", "#NAME?");
            RunOK("B9 empty unquoted", cp => cp.Value = "", "", "0.0000");
            RunOK("B10 null Value (cell unwritten, Visio default)", cp => cp.Value = (string)null, "0", "0.0000");

            if (failures.Count > 0)
            {
                MUT.Assert.Fail("\n" + string.Join("\n", failures));
            }

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }

        [MUT.TestMethod]
        public void CustomProps_DateTypeCharacterization()
        {
            // Type=Date (Type=5). Behavior matrix (locked 2026-05-06):
            //
            // Input                                   | Outcome
            // ----------------------------------------+-----------------------------------------------------
            // DATETIME(\"03/31/2017 14:05:06\")       | succeeds, formula round-trips, result=3/31/2017 2:05:06 PM (locale-formatted)
            // "testVal" plain identifier              | THROWS COMException #NAME?
            // "\"2017-03-31\"" pre-quoted ISO date    | succeeds as a literal string, formula=\"2017-03-31\", result=2017-03-31 (NOT parsed as a date — Type metadata mismatch)
            // ""    empty                             | succeeds, formula=[empty], result=0.0000
            // null                                    | HasValue=false, cell unwritten; default formula=0, result=0.0000

            var page1 = this.GetNewPage();
            var failures = new System.Collections.Generic.List<string>();
            int caseIndex = 0;

            string MakePropName() { caseIndex++; return "D" + caseIndex; }

            void RunOK(string label, System.Action<CustomPropertyCells> setup, string expFormula, string expResult)
            {
                string propName = MakePropName();
                var s = page1.DrawRectangle(0, 0, 1, 1);
                var cp = new CustomPropertyCells();
                cp.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Date);
                try
                {
                    setup(cp);
                    VA.Shapes.CustomPropertyHelper.Set(s, propName, cp);
                    var fdic = VA.Shapes.CustomPropertyHelper.GetDictionary(s, VisioAutomation.Core.CellValueType.Formula);
                    var rdic = VA.Shapes.CustomPropertyHelper.GetDictionary(s, VisioAutomation.Core.CellValueType.Result);
                    string af = fdic.ContainsKey(propName) ? (fdic[propName].Value.Value ?? "<null>") : "<missing>";
                    string ar = rdic.ContainsKey(propName) ? (rdic[propName].Value.Value ?? "<null>") : "<missing>";
                    if (af != expFormula || ar != expResult)
                    {
                        failures.Add(string.Format("[{0}] exp formula=[{1}] result=[{2}], got formula=[{3}] result=[{4}]",
                            label, expFormula, expResult, af, ar));
                    }
                }
                catch (System.Exception ex)
                {
                    failures.Add(string.Format("[{0}] expected success but THREW {1}: {2}", label, ex.GetType().Name, ex.Message));
                }
            }

            void RunThrows(string label, System.Action<CustomPropertyCells> setup, string expMsgContains)
            {
                string propName = MakePropName();
                var s = page1.DrawRectangle(0, 0, 1, 1);
                var cp = new CustomPropertyCells();
                cp.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Date);
                try
                {
                    setup(cp);
                    VA.Shapes.CustomPropertyHelper.Set(s, propName, cp);
                    failures.Add(string.Format("[{0}] expected COMException with [{1}] but Set succeeded", label, expMsgContains));
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    if (!ex.Message.Contains(expMsgContains))
                    {
                        failures.Add(string.Format("[{0}] threw COMException but message [{1}] doesn't contain [{2}]", label, ex.Message, expMsgContains));
                    }
                }
                catch (System.Exception ex)
                {
                    failures.Add(string.Format("[{0}] expected COMException but threw {1}: {2}", label, ex.GetType().Name, ex.Message));
                }
            }

            RunOK("D1 DATETIME formula", cp => cp.Value = "DATETIME(\"03/31/2017 14:05:06\")", "DATETIME(\"03/31/2017 14:05:06\")", "3/31/2017 2:05:06 PM");
            RunThrows("D2 plain identifier", cp => cp.Value = "testVal", "#NAME?");
            RunOK("D3 quoted ISO date (stored as literal string, Type metadata mismatch)", cp => cp.Value = "\"2017-03-31\"", "\"2017-03-31\"", "2017-03-31");
            RunOK("D4 empty unquoted", cp => cp.Value = "", "", "0.0000");
            RunOK("D5 null Value (cell unwritten, Visio default)", cp => cp.Value = (string)null, "0", "0.0000");

            if (failures.Count > 0)
            {
                MUT.Assert.Fail("\n" + string.Join("\n", failures));
            }

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }
    }
}