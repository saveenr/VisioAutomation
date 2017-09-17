using System;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace VisioAutomation_Tests.Core.ShapeSheet
{
    [TestClass]
    public class CellGroupTests : VisioAutomationTest
    {
        [TestMethod]
        public void VerifyAllCellsAreEnumerated()
        {
            var types = new List<Type>();
            types.Add(typeof(VisioAutomation.Shapes.ControlCells));
            types.Add(typeof(VisioAutomation.Shapes.ConnectionPointCells));
            types.Add(typeof(VisioAutomation.Shapes.CustomPropertyCells));
            types.Add(typeof(VisioAutomation.Shapes.HyperlinkCells));
            types.Add(typeof(VisioAutomation.Shapes.LockCells));
            types.Add(typeof(VisioAutomation.Shapes.ShapeFormatCells));
            types.Add(typeof(VisioAutomation.Shapes.ShapeLayoutCells));
            types.Add(typeof(VisioAutomation.Shapes.ShapeXFormCells));
            types.Add(typeof(VisioAutomation.Pages.PageFormatCells));
            types.Add(typeof(VisioAutomation.Pages.PageLayoutCells));
            types.Add(typeof(VisioAutomation.Pages.PagePrintCells));
            types.Add(typeof(VisioAutomation.Pages.PageRulerAndGridCells));

            foreach (var type in types)
            {
                string type_name = type.Name;
                var ctor = type.GetConstructor(Type.EmptyTypes);
                var type_obj = ctor.Invoke(new object[] { });
                var cellgroup = (VisioAutomation.ShapeSheet.CellGroups.CellGroupBase) type_obj;

                var props = GetCellDataProps(type);
                var reflected_cvts = props.Select(p => (VisioAutomation.ShapeSheet.CellValueLiteral)p.GetValue(cellgroup, null)).ToList();
                var reflected_cvt_values = reflected_cvts.Select(p => p.Value).ToList();
                var reflected_cvt_names = props.Select(p => p.Name).ToList();
                var reflected_nametovalue = new Dictionary<string,string>();
                for (int i = 0; i < props.Count; i++)
                {
                    string k = reflected_cvt_names[i];
                    string v = reflected_cvt_values[i];
                    reflected_nametovalue[k] = v;
                }

                var enumerated_values = cellgroup.SrcValuePairs.Select(i => i.Value).ToList();
                var enumerated_srcs = cellgroup.SrcValuePairs.Select(i => i.Src).ToList();
                var enumerated_srctovalue = cellgroup.SrcValuePairs.ToDictionary(i => i.Src, i => i.Value);

                Assert.AreEqual(reflected_cvts.Count, enumerated_values.Count);
            }
        }

        private static List<PropertyInfo> GetCellDataProps(Type t)
        {
            var props = t.GetProperties().Where(p => p.MemberType == MemberTypes.Property).ToList();
            var cellprops = props.Where(p => p.PropertyType == typeof(VisioAutomation.ShapeSheet.CellValueLiteral)).ToList();
            return cellprops;
        }


    }
}