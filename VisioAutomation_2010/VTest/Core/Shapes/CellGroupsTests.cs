using System.Linq;
using System.Reflection;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace VTest.Core.ShapeSheet
{
    [MUT.TestClass]
    public class CellGroupTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void VerifyAllCellsAreEnumerated()
        {
            var types = new List<System.Type>();
            types.Add(typeof(VisioAutomation.Shapes.ControlCells));
            types.Add(typeof(VisioAutomation.Shapes.ConnectionPointCells));
            types.Add(typeof(VisioAutomation.Shapes.CustomPropertyCells));
            types.Add(typeof(VisioAutomation.Shapes.HyperlinkCells));
            types.Add(typeof(VisioAutomation.Shapes.LockCells));
            types.Add(typeof(VisioAutomation.Shapes.FormatCells));
            types.Add(typeof(VisioAutomation.Shapes.LayoutCells));
            types.Add(typeof(VisioAutomation.Shapes.XFormCells));
            types.Add(typeof(VisioAutomation.Pages.FormatCells));
            types.Add(typeof(VisioAutomation.Pages.LayoutCells));
            types.Add(typeof(VisioAutomation.Pages.PrintCells));
            types.Add(typeof(VisioAutomation.Pages.RulerAndGridCells));

            var cvt_ctor = typeof(VisioAutomation.Core.CellValue).GetConstructor(new []{typeof(string)});
            foreach (var cellgroup_type in types)
            {
                var cellgroup_ctor = cellgroup_type.GetConstructor(System.Type.EmptyTypes);
                var cellgroup_obj = cellgroup_ctor.Invoke(new object[] { });
                var cellgroup = (VisioAutomation.ShapeSheet.CellGroups.CellRecord) cellgroup_obj;

                var props = _get_cell_data_props(cellgroup_type);

                // Set unique values for the cells
                // Later we'll verify they can be retrieved

                var input_values = Enumerable.Range(0, props.Count).Select(i => i.ToString()).ToList();
                for (int i = 0; i < props.Count; i++)
                {
                    var prop = props[i];
                    var cvl_value = cvt_ctor.Invoke(new object[] {input_values[i]});
                    prop.SetValue(cellgroup, cvl_value);
                }

                var reflected_cvts = props.Select(p => (VisioAutomation.Core.CellValue)p.GetValue(cellgroup, null)).ToList();
                var reflected_cvt_values = reflected_cvts.Select(p => p.Value).ToList();
                var reflected_cvt_names = props.Select(p => p.Name).ToList();
                var reflected_nametovalue = new Dictionary<string,string>();
                for (int i = 0; i < props.Count; i++)
                {
                    string k = reflected_cvt_names[i];
                    string v = reflected_cvt_values[i];
                    reflected_nametovalue[k] = v;
                }

                var enumerated_values = cellgroup.GetSrcValues().Select(i => i.Value).ToList();
                var enumerated_srcs = cellgroup.GetSrcValues().Select(i => i.Src).ToList();
                var enumerated_srctovalue = cellgroup.GetSrcValues().ToDictionary(i => i.Src, i => i.Value);

                MUT.Assert.AreEqual(reflected_cvts.Count, enumerated_values.Count);

                // Verify that all the enumerated Srcs are distinct
                var unique_enumerated_srcs = enumerated_srcs.Distinct().ToList();
                MUT.Assert.AreEqual(enumerated_srcs.Count, unique_enumerated_srcs.Count);

                // Verify that all the enumerated values are distinct
                var unique_enumerated_values = enumerated_values.Distinct().ToList();
                MUT.Assert.AreEqual(reflected_cvts.Count, unique_enumerated_values.Count);

                foreach (var input_value in input_values)
                {
                    //MUT.Assert.IsTrue(unique_enumerated_values.Contains(input_value));
                }

            }
        }

        private static List<PropertyInfo> _get_cell_data_props(System.Type t)
        {
            var props = t.GetProperties().Where(p => p.MemberType == MemberTypes.Property).ToList();
            var cellprops = props.Where(p => p.PropertyType == typeof(VisioAutomation.Core.CellValue)).ToList();
            return cellprops;
        }


    }
}