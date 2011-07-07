using System.Collections.Generic;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using VisioAutomation.Metadata;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class MetadataTests : VisioAutomationTest
    {
        [TestMethod]
        public void CheckPersistance()
        {
            string output_path = VisioTestCommon.Helper.GetTestMethodOutputFilename();

            if (!System.IO.Directory.Exists(output_path))
            {
                System.IO.Directory.CreateDirectory(output_path);
            }
            var db = VA.Metadata.MetadataDB.Load();
            db.Save(output_path);
        }

        [TestMethod]
        public void VerifyMetadaDBCreation()
        {
            var db = VA.Metadata.MetadataDB.Load();

            var allcells = db.Cells;

            var dupe_cell_names = TestHelper.GetDuplicates(allcells.Select(c => c.Name));
            Assert.IsTrue(dupe_cell_names.Contains("Tabs"));

            Assert.AreEqual(346, allcells.Count);

            var visio_2007_cells = allcells.Where(c => c.MinVersion.Contains("Visio2007")).ToList();
            Assert.AreEqual(343, visio_2007_cells.Count());

            TestHelper.AssertNoDuplicates(allcells.Select(c => c.ID));
        }

        [TestMethod]
        public void Constants()
        {
            var db = VA.Metadata.MetadataDB.Load();

            var constants = db.Constants;

            // There are 3003 known constants in the Visio PIA
            Assert.AreEqual(3003, constants.Count);
        }
        
        [TestMethod]
        public void Sections()
        {
            var db = VA.Metadata.MetadataDB.Load();

            var sections = db.Sections;

            // There are 40 known sections in the Visio PIA
            Assert.AreEqual(40, sections.Count);
        }

        [TestMethod]
        public void CellValues()
        {
            var db = VA.Metadata.MetadataDB.Load();
            var cellvals = db.CellValues;

            // There are 40 known sections in the Visio PIA
            Assert.AreEqual(397, cellvals.Count);
        }

        [TestMethod]
        public void ValidateCellNameCode()
        {
            var db = VA.Metadata.MetadataDB.Load();

            var cellvals = db.CellValues;
            var allcells = db.Cells;
            var visio_2007_cells = allcells.Where(c => c.MinVersion.Contains("Visio2007")).ToList();

            var va_name_to_src = VA.ShapeSheet.SRCConstants.GetSRCDictionary();

            TestHelper.AssertNoDuplicates(visio_2007_cells.Select(i => i.NameCode));

            var db_codename_to_cell = visio_2007_cells.ToDictionary(i => i.NameCode, i => i);

            var unfound = new List<string>();

            // verify that all the fields in SRCConstants are corrected represented in the metadata
            foreach (var va_name in va_name_to_src.Keys)
            {
                if (!db_codename_to_cell.ContainsKey(va_name))
                {
                    unfound.Add(va_name);
                }
            }

            if (unfound.Count > 0)
            {
                string message = string.Format(" didn't find in DB cells " + string.Join(",", unfound));
                
            }

            unfound.Clear();
            // verify that all the fields in db are corrected represented in the VA srcconstants
            foreach (var db_name in db_codename_to_cell.Keys)
            {
                if (!va_name_to_src.ContainsKey(db_name))
                {
                    unfound.Add(db_name);
                }
            }

            if (unfound.Count > 0)
            {
                string message = string.Format(" didn't find in src constants " + string.Join(",", unfound));

            }


            int x = 1;
        }

        [TestMethod]
        public void CheckPIA()
        {
            var db = VA.Metadata.MetadataDB.Load();
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
                var pia_enum_values = TestHelper.EnumToDictionary<int>(pia_enum);
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
                var pia_dic = TestHelper.EnumToDictionary<int>(pia_type);
                foreach (string md_value_name in md_enum.Items.Select(i=>i.Name))
                {

                    Assert.IsTrue(pia_dic.ContainsKey(md_value_name));
                    Assert.AreEqual(md_enum[md_value_name],pia_dic[md_value_name]);
                }
            }

        }

        [TestMethod]
        public void CheckSRCConstantIndices()
        {
            var db = VA.Metadata.MetadataDB.Load();
            var all_cells = db.Cells;
            var visio_2007_cells = all_cells.Where(c => c.MinVersion.Contains("Visio2007")).ToList();
            var va_name_to_src = VA.ShapeSheet.SRCConstants.GetSRCDictionary();
            var db_name_to_cell = visio_2007_cells.ToDictionary(c => c.NameCode, c => c);
            foreach (string name in va_name_to_src.Keys)
            {
                if (!db_name_to_cell.ContainsKey(name))
                {
                    Assert.Fail("DB does not contain sll with namecode " + name);
                }
                
            }

            var sectioindexname_to_int = TestHelper.EnumToDictionary<int>(typeof(IVisio.VisSectionIndices));
            var rowindexname_to_int = TestHelper.EnumToDictionary<int>(typeof(IVisio.VisRowIndices));
            var cellindexname_to_int = TestHelper.EnumToDictionary<int>(typeof(IVisio.VisCellIndices));
            foreach (var db_cell in visio_2007_cells)
            {
                if (!sectioindexname_to_int.ContainsKey(db_cell.SectionIndex))
                {
                    Assert.Fail(db_cell.Name);
                }
                if (!rowindexname_to_int.ContainsKey(db_cell.RowIndex))
                {
                    Assert.Fail(db_cell.Name);
                }
                if (!cellindexname_to_int.ContainsKey(db_cell.CellIndex))
                {
                    Assert.Fail(db_cell.CellIndex + " " + db_cell.Name);
                }

                int s = sectioindexname_to_int[db_cell.SectionIndex];
                int r = rowindexname_to_int[db_cell.RowIndex];
                int c = cellindexname_to_int[db_cell.CellIndex];

                if (!va_name_to_src.ContainsKey(db_cell.NameCode))
                {
                    Assert.Fail(db_cell.NameCode);
                }
            }
        }



        [TestMethod]
        public void CheckSectionIndices()
        {
            var db = VA.Metadata.MetadataDB.Load();

            // verify that each section has an sectioindex enum that is found in the database
            foreach (var section in db.Sections)
            {
                string secindex_name = section.Enum;
                int secindex_int = db.GetAutomationConstantByName(secindex_name).GetValueAsInt();
            }
           

        }



        [TestMethod]
        public void CheckDBCellNames()
        {
            var app = new IVisio.Application();
            var docs = app.Documents;
            var doc = docs.Add("");
            var page = doc.Pages[1];

            var shape1 = page.DrawRectangle(2,2 ,5,6);
            shape1.Text = "0123456789\n" + "0123456789\n" + "0123456789\n" + "0123456789\n" + "0123456789\n" + "0123456789\n" +
                          "0123456789\n";
            var fmt1 = new VA.Text.CharacterFormatCells();
            fmt1.Color = 3;
            VA.Text.TextHelper.SetFormat( shape1, fmt1, 10 , 20);
            VA.Text.TextHelper.SetFormat( shape1, fmt1, 35, 45);

            var fmt2 = new VA.Text.ParagraphFormatCells();
            VA.Text.TextHelper.SetFormat(shape1, fmt2, 30, 40);

            var cp1 = new VA.Connections.ConnectionPointCells();
            cp1.X = "Width";
            cp1.Y = "Height*0.5";

            VA.Connections.ConnectionPointHelper.AddConnectionPoint(shape1, cp1);

            var db = VA.Metadata.MetadataDB.Load();
            var all_cells = db.Cells;
            var visio_2007_cells = all_cells.Where(c => c.MinVersion.Contains("Visio2007")).ToList();

            var data = new[]
                           {
                               new {shape = shape1, sec=IVisio.VisSectionIndices.visSectionObject, obj="shape"},
                               new {shape = shape1, sec=IVisio.VisSectionIndices.visSectionCharacter, obj="shape"},
                               new {shape = shape1, sec=IVisio.VisSectionIndices.visSectionParagraph, obj="shape"},
                               new {shape = page.PageSheet, sec=IVisio.VisSectionIndices.visSectionObject, obj="page"},
                               new {shape = shape1, sec=IVisio.VisSectionIndices.visSectionConnectionPts, obj="shape"}
                           };


            foreach (var datum in data)
            {
                var shape = datum.shape;
                var secobj = db.GetSectionBySectionIndex((int)datum.sec);
                var target_cells = visio_2007_cells.Where(c => c.SectionIndex == secobj.Enum).Where(c => c.Object.Contains(datum.obj)).ToList();

                foreach (var db_cell in target_cells)
                {
                    var md_section = db.GetAutomationConstantByName(db_cell.SectionIndex);
                    var md_row = db.GetAutomationConstantByName(db_cell.RowIndex);
                    var md_cell = db.GetAutomationConstantByName(db_cell.CellIndex);

                    var s = (short)md_section.GetValueAsInt();
                    var r = (short)md_row.GetValueAsInt();
                    var c = (short)md_cell.GetValueAsInt();
                    var src = new VA.ShapeSheet.SRC(s, r, c);

                    // Verify that the VisioAutomation library can find this cell
                    var va_cellname = VA.ShapeSheet.ShapeSheetHelper.TryGetNameFromSRC(src);
                    if (va_cellname == null)
                    {
                        string msg = string.Format(@" DB Cell not found in VisioAutomation: ""{0}"" ({1},{2},{3}) ",
                                                   db_cell.Name,
                                                   md_section.Name,
                                                   md_row.Name,
                                                   md_cell.Name
                            );
                        Assert.Fail(msg);

                    }

                    // Verify that the Visio application can find this cell naame
                    var pia_cell = shape.CellsSRC[s, r, c];
                    string piacellname = pia_cell.Name;
                    if (db_cell.Name != piacellname)
                    {
                        if (r != (short)IVisio.VisRowIndices.visRow1stHyperlink)
                        {
                            string msg0 = string.Format("Names don't match. DB Cell Name =\"{0}\" but Actual Cell Name = \"{1}\"",
                                                       db_cell.Name, piacellname);

                            Assert.Fail(msg0);

                        }
                    }
                }

            }

            /// end
            app.Quit(true);

        }

        [TestMethod]
        public void ExportMetadataCode()
        {
            var db = VA.Metadata.MetadataDB.Load();

            string filename = VisioTestCommon.Helper.GetTestMethodOutputFilename(".txt");

            db.ExportCode(filename);

        }
    }
}