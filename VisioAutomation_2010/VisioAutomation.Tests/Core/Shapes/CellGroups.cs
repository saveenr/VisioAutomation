using System;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes;
using VisioAutomation.Extensions;
using System.Collections.Generic;

namespace VisioAutomation_Tests.Core.Shapes
{
    [TestClass]
    public class CellGroupTests : VisioAutomationTest
    {
        [TestMethod]
        public void EnumCellgroupCells()
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

            var xg1 = new VisioAutomation.Shapes.ShapeXFormCells();
            xg1.PinX = 1.0;
            xg1.PinY = 2.0;

            var xg1_type = xg1.GetType();
            var props = GetCellDataProps(xg1_type);

            var cellvalues = props.Select(p => (VisioAutomation.ShapeSheet.CellData)p.GetValue(xg1,null)).ToList();
            var cellvalues_formulas = cellvalues.Select(p=>p.Value).ToList();

            var cellnames = props.Select(p => p.Name).ToList();

            var f2 = xg1.SrcFormulaPairs.Select(i => i.Formula).ToList();
            
            int x = 1;
        }

        private static List<PropertyInfo> GetCellDataProps(Type t)
        {
            var props = t.GetProperties().Where(p => p.MemberType == MemberTypes.Property).ToList();
            var cellprops = props.Where(p => p.PropertyType == typeof(VisioAutomation.ShapeSheet.CellData)).ToList();
            return cellprops;
        }


    }
}