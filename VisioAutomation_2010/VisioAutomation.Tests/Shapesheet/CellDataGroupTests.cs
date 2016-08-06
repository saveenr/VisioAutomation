using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{

    class G1
    {
        int x;
        public int foo { get; set; }
        public int? bar { get; set; }
        public VA.ShapeSheet.CellData<double> A { get; set; }
        public VA.ShapeSheet.CellData<int> B { get; set; }
        public VA.ShapeSheet.CellData<bool> C { get; set; }
        public VA.ShapeSheet.CellData<string> D { get; set; }
    }

    public class Rec
    {
        public string Name;
        public int Ordinal;
        public System.Type StorageType;
        public System.Reflection.PropertyInfo PropInfo;
    }

    [TestClass]
    public class CellDataGroupTests : VisioAutomationTest
    {
        public static bool IsAssignableToGenericType(System.Type givenType, System.Type genericType)
        {
            var interfaceTypes = givenType.GetInterfaces();

            foreach (var it in interfaceTypes)
                if (it.IsGenericType)
                    if (it.GetGenericTypeDefinition() == genericType) return true;

            System.Type baseType = givenType.BaseType;
            if (baseType == null) return false;

            return baseType.IsGenericType &&
                baseType.GetGenericTypeDefinition() == genericType ||
                IsAssignableToGenericType(baseType, genericType);
        }

        public static bool FFF(System.Type t)
        {
            if ( typeof(VA.ShapeSheet.CellData<double>).IsAssignableFrom(t))
            {
                return true;
            }
            if ( typeof(VA.ShapeSheet.CellData<int>).IsAssignableFrom(t))
            {
                return true;
            }
            if ( typeof(VA.ShapeSheet.CellData<bool>).IsAssignableFrom(t))
            {
                return true;
            }
            if ( typeof(VA.ShapeSheet.CellData<string>).IsAssignableFrom(t))
            {
                return true;
            }
            return false;
        }

        public static IList<Rec> getcelldataprops(System.Type t)
        {
            var all_props = typeof(G1).GetProperties();
            var gen_props = all_props.Where(p => p.PropertyType.IsGenericType).ToArray();
            var cd_props = gen_props.Where(p => FFF(p.PropertyType)).ToArray();

            var recs = new List<Rec>();
            foreach (var cd_prop in cd_props)
            {
                var rec = new Rec();
                rec.Name = cd_prop.Name;
                rec.Ordinal = recs.Count();
                rec.StorageType = cd_prop.PropertyType.GetGenericArguments()[0];
                rec.PropInfo = cd_prop;

                recs.Add(rec);
            }
            return recs;

        }
        [TestMethod]
        public void T1()
        {
            var cd_props = getcelldataprops(typeof (G1));
            Assert.AreEqual(4, cd_props.Count());

            var x1 = new G1();

            var q = new VA.ShapeSheet.Query.CellQuery();
            foreach (var cd_prop in cd_props)
            {
                q.AddColumn()
            }
        }
    }
}
