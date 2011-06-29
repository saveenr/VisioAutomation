using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class MetadataTests : VisioAutomationTest
    {
        [TestMethod]
        public void VerifyMetadaDBCreation()
        {
            var db = new VA.Metadata.MetadataDB();

            var allcells = db.Cells;

            var dupe_cell_names = get_dupes(allcells.Select(c => c.Name));
            Assert.IsTrue(dupe_cell_names.Contains("Tabs"));
            Assert.IsTrue(dupe_cell_names.Contains("HideForApply"));

            Assert.AreEqual(346, allcells.Count);

            var visio_2007_cells = allcells.Where(c => c.MinVersion.Contains("Visio2007")).ToList();
            Assert.AreEqual(344, visio_2007_cells.Count());

            no_dupes(allcells.Select(c => c.ID));
        }

        [TestMethod]
        public void Constants()
        {
            var db = new VA.Metadata.MetadataDB();

            var md_constants = db.Constants;

            // There are 3003 known constants in the Visio PIA
            Assert.AreEqual(3003, md_constants.Count);

            //Check that each constant maps to what is actually in the Visio PIO

            var pia_enum_types = VisioAutomation.Interop.InteropHelper.GetEnumTypes();

            // Visio 2007 has 116 enum types
            // Visio 2010 has ??? enum types

            var pia_name_to_enum = pia_enum_types.ToDictionary(i => i.Name, i => i);
            var md_names = new HashSet<string>(md_constants.Select(i => i.Enum));

            foreach (var md_constant in md_constants)
            {
                if (!pia_name_to_enum.ContainsKey(md_constant.Enum))
                {
                    Assert.Fail("metadata is missing a PIA enum");
                }
            }

            foreach (var enum_type in pia_enum_types)
            {
                if (!md_names.Contains(enum_type.Name))
                {
                    Assert.Fail("Metadata has a enum that is not in the PIA");
                }
            }

            var md_enums = db.AutomationEnums;

            foreach (var md_enum in md_enums)
            {
                var pia_enum = pia_name_to_enum[md_enum.Name];
                var pia_value_dic = GetEnumValues<int>(pia_enum);

                foreach (string pia_vname in pia_value_dic.Keys)
                {
                    if (!md_enum.HasItem(pia_vname))
                    {
                        Assert.Fail("DB missing enum value");                        
                    }
                }

                foreach (string md_vname in md_enum.Items.Select(i=>i.Name))
                {
                    if (!pia_value_dic.ContainsKey(md_vname))
                    {
                        Assert.Fail("DB has a value that PIA does not");
                    }
                }

            }

            

            int x = 1;

        }

        public IDictionary<string,T> GetEnumValues<T>(System.Type t )
        {
            var values = System.Enum.GetValues(t);
            var names = System.Enum.GetNames(t);
            var dic = new Dictionary<string, T>(names.Length);
            var typed_values = new T[values.Length];
            for (int i = 0; i < typed_values.Length; i++)
            {
                dic[names[i]] = (T) values.GetValue(i);
            }

            return dic;
        }


        [TestMethod]
        public void Sections()
        {
            var db = new VA.Metadata.MetadataDB();

            var sections = db.Sections;

            // There are 40 known sections in the Visio PIA
            Assert.AreEqual(40, sections.Count);
        }

        [TestMethod]
        public void CellValues()
        {
            var db = new VA.Metadata.MetadataDB();

            var cellvals = db.CellValues;

            // There are 40 known sections in the Visio PIA
            Assert.AreEqual(397, cellvals.Count);
        }

        [TestMethod]
        public void ValidateSRCConstants()
        {
            var db = new VA.Metadata.MetadataDB();

            var cellvals = db.CellValues;

            var allcells = db.Cells;
            var visio_2007_cells = allcells.Where(c => c.MinVersion.Contains("Visio2007")).ToList();

            var fields_name_to_value = this.GetSRCDictionary();

            
            foreach (var src_field_name in fields_name_to_value.Keys)
            {
                var src = fields_name_to_value[src_field_name];

                var src_cellname = VA.ShapeSheet.ShapeSheetHelper.TryGetNameFromSRC(src);
            }

            int x = 1;
        }

        public Dictionary<string, VA.ShapeSheet.SRC> GetSRCDictionary()
        {
            var fields = GetSRCFields();

            var fields_name_to_value = new Dictionary<string, VA.ShapeSheet.SRC>();
            foreach (var field in fields)
            {
                fields_name_to_value[field.Name] = (VA.ShapeSheet.SRC)field.GetValue(null);
            }

            return fields_name_to_value;
        }

        public List<FieldInfo> GetSRCFields()
        {
            var srcconstants_t = typeof (VA.ShapeSheet.SRCConstants);
            var fields = srcconstants_t.GetFields()
                .Where(m => m.FieldType == typeof (VA.ShapeSheet.SRC))
                .Where(m => m.IsPublic)
                .Where(m => m.IsStatic)
                .ToList();
            return fields;
        }

        public List<T> get_dupes<T>(IEnumerable<T> items)
        {
            var set = new HashSet<T>();
            var dupes = new List<T>();

            foreach (var item in items)
            {
                if (set.Contains(item))
                {
                    dupes.Add(item);
                }
                else
                {
                    set.Add(item);
                }
            }

            return dupes;
        }

        public void no_dupes<T>(IEnumerable<T> items)
        {
            var set = new HashSet<T>();
            var dupes = new List<T>();

            foreach (var item in items)
            {
                if (set.Contains(item))
                {
                    dupes.Add(item);
                }
                else
                {
                    set.Add(item);
                }
            }

            if (dupes.Count > 0)
            {
                Assert.Fail(string.Format("Duplicated {0}", dupes.Count));
            }
        }


    }
}