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

            var constants = db.Constants;

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

            var fields_name_to_value = VA.ShapeSheet.SRCConstants.GetSRCDictionary();
            
            foreach (var src_field_name in fields_name_to_value.Keys)
            {
                var src = fields_name_to_value[src_field_name];
                var src_cellname = VA.ShapeSheet.ShapeSheetHelper.TryGetNameFromSRC(src);
            }

            int x = 1;
        }

        [TestMethod]
        public void CheckPIA()
        {
            var db = new VA.Metadata.MetadataDB();
            var db_autoenums = db.AutomationEnums;

            var pia_enums = VA.Interop.InteropHelper.GetEnumTypes();

            var db_name_to_enum = db_autoenums.ToDictionary(i => i.Name, i=>i);
            foreach (var pia_enum in pia_enums)
            {


                Assert.IsTrue( db_name_to_enum.ContainsKey(pia_enum.Name));
            }

            // verify that everying in the metadatadb is int the PIA 

            foreach (var pia_enum in pia_enums)
            {
                var pia_enum_values = GetNameToValueMap<int>(pia_enum);
                var db_enum = db.GetAutomationEnumByName(pia_enum.Name);
                foreach (string pia_value_name in pia_enum_values.Keys)
                {
                    Assert.IsTrue(db_enum.HasItem(pia_value_name));
                    Assert.AreEqual(pia_enum_values[pia_value_name],db_enum[pia_value_name]);
                }
            }


            // verify that everying in the PIA is int the metadatadb

            var name_to_pia_type = pia_enums.ToDictionary(i => i.Name, i => i);

            foreach (var md_enum  in db.AutomationEnums)
            {
                var pia_type = name_to_pia_type[md_enum.Name];
                var pia_dic = GetNameToValueMap<int>(pia_type);
                foreach (string md_value_name in md_enum.Items.Select(i=>i.Name))
                {

                    Assert.IsTrue(pia_dic.ContainsKey(md_value_name));
                    Assert.AreEqual(md_enum[md_value_name],pia_dic[md_value_name]);
                }
            }

        }

        public Dictionary<string,T> GetNameToValueMap<T>( System.Type t)
        {
            var dic = new Dictionary<string, T>();
            string[] names = System.Enum.GetNames(t);
            System.Array avalues = System.Enum.GetValues(t);
            for (int i = 0; i < avalues.Length; i++)
            {
                
                dic[names[i]] = (T)avalues.GetValue(i);
            }

            return dic;

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