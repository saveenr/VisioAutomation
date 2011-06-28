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

            var constants = db.AutomationEnums;

            // There are 3003 known constants in the Visio PIA
            Assert.AreEqual(3003, constants.Count);
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