using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Core.ShapeSheet
{
    [TestClass]
    public class ShapeSheetQueryTests : VisioAutomationTest
    {
        public static VA.ShapeSheet.Src cell_fg = VA.ShapeSheet.SrcConstants.FillForeground;
        public static VA.ShapeSheet.Src cell_bg = VA.ShapeSheet.SrcConstants.FillBackground;
        public static VA.ShapeSheet.Src cell_pat = VA.ShapeSheet.SrcConstants.FillPattern;

        [TestMethod]
        public void ShapeSheet_Query_SectionCells_have_names()
        {
            var query = new SectionsQuery();

            var sec_char = query.AddSubQuery(IVisio.VisSectionIndices.visSectionCharacter);
            Assert.AreEqual("Character", sec_char.Name);

            var sec_obj = query.AddSubQuery(IVisio.VisSectionIndices.visSectionObject);
            Assert.AreEqual("Object", sec_obj.Name);

        }

        [TestMethod]
        public void ShapeSheet_Query_GetResults_SingleShape()
        {
            var doc1 = this.GetNewDoc();
            var page1 = doc1.Pages[1];
            VisioAutomationTest.SetPageSize(page1, this.StandardPageSize);

            // draw a simple shape
            var s1 = page1.DrawRectangle(this.StandardPageSizeRect);
            int s1_id = s1.ID;

            // format it with setformulas
            var fg_cell = s1.Cells["FillForegnd"];
            var bg_cell = s1.Cells["FillBkgnd"];
            var pat_cell = s1.Cells["FillPattern"];

            fg_cell.FormulaU = "RGB(255,0,0)";
            bg_cell.FormulaU = "RGB(0,0,255)";
            pat_cell.FormulaU = "40";

            // now retrieve the formulas with GetFormulas

            var src_fg = VA.ShapeSheet.SrcConstants.FillForeground;
            var src_bg = VA.ShapeSheet.SrcConstants.FillBackground;
            var src_filpat = VA.ShapeSheet.SrcConstants.FillPattern;

            var query = new CellQuery();
            var col_fg = query.Columns.Add(src_fg, "FillForegnd");
            var col_bg = query.Columns.Add(src_bg, "FillBkgnd");
            var col_filpat = query.Columns.Add(src_filpat, "FillPattern");

            var shapeids = new[] {s1_id};

            var formulas = query.GetFormulas(page1, shapeids);

            // now verify that the formulas were actually set
            Assert.AreEqual("RGB(255,0,0)", formulas[0].Cells[col_fg]);
            Assert.AreEqual("RGB(0,0,255)", formulas[0].Cells[col_bg]);
            Assert.AreEqual("40", formulas[0].Cells[col_filpat]);

            // now retrieve the results with GetResults as floats
            var float_results = query.GetResults<double>(page1,shapeids);
            Assert.IsNotNull(float_results);
            Assert.AreEqual(40.0, float_results[0].Cells[col_filpat]);

            // now retrieve the results with GetResults as ints
            var int_results = query.GetResults<int>(page1, shapeids);
            Assert.AreEqual(40, int_results[0].Cells[col_filpat]);

            // now retrieve the results with GetResults as strings

            var string_results = query.GetResults<string>(page1, shapeids);

            Assert.AreEqual("RGB(255, 0, 0)", string_results[0].Cells[col_fg]);
            Assert.AreEqual("RGB(0, 0, 255)", string_results[0].Cells[col_bg]);
            Assert.AreEqual("40", string_results[0].Cells[col_filpat]);

            page1.Delete(0);
            doc1.Close(true);
        }

        [TestMethod]
        public void ShapeSheet_Query_GetResults_MultipleShapes()
        {
            var page1 = this.GetNewPage();

            // draw a simple shape
            var s1 = page1.DrawRectangle(this.StandardPageSizeRect);
            int s1_id = s1.ID;

            // format it with setformulas
            var fg_cell = s1.Cells["FillForegnd"];
            var bg_cell = s1.Cells["FillBkgnd"];
            var pat_cell = s1.Cells["FillPattern"];

            fg_cell.ResultIU = 2.0; //red
            bg_cell.ResultIU = 3.0; //green
            pat_cell.ResultIU = 40.0;

            var src_fg = VA.ShapeSheet.SrcConstants.FillForeground;
            var src_bg = VA.ShapeSheet.SrcConstants.FillBackground;
            var src_filpat = VA.ShapeSheet.SrcConstants.FillPattern;

            // now retrieve the formulas with GetFormulas

            var query = new CellQuery();
            var col_fg = query.Columns.Add(src_fg, "FillForegnd");
            var col_bg = query.Columns.Add(src_bg, "FillBkgnd");
            var col_filpat = query.Columns.Add(src_filpat, "FillPattern");

            var shapeids = new[] {s1_id};

            var formulas = query.GetFormulas(page1, shapeids);

            // now verify that the formulas were actually set
            Assert.AreEqual("2", formulas[0].Cells[col_fg]);
            Assert.AreEqual("3", formulas[0].Cells[col_bg]);
            Assert.AreEqual("40", formulas[0].Cells[col_filpat]);

            // now retrieve the results with GetResults as floats
            var float_results = query.GetResults<double>(page1, shapeids);
            Assert.AreEqual(2.0, float_results[0].Cells[col_fg]);
            Assert.AreEqual(3.0, float_results[0].Cells[col_bg]);
            Assert.AreEqual(40.0, float_results[0].Cells[col_filpat]);

            // now retrieve the results with GetResults as ints
            var int_results = query.GetResults<int>(page1, shapeids);

            Assert.AreEqual(2, int_results[0].Cells[col_fg]);
            Assert.AreEqual(3, int_results[0].Cells[col_bg]);
            Assert.AreEqual(40, int_results[0].Cells[col_filpat]);

            // now retrieve the results with GetResults as strings
            var string_results = query.GetResults<string>(page1,shapeids);
            Assert.AreEqual("2", string_results[0].Cells[col_fg]);
            Assert.AreEqual("3", string_results[0].Cells[col_bg]);
            Assert.AreEqual("40", string_results[0].Cells[col_filpat]);

            page1.Delete(0);
        }

        [TestMethod]
        public void ShapeSheet_Query_SectionRowHandling()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(2, 1, 3, 3);
            var s3 = page1.DrawRectangle(3, 1, 4, 2);
            var s4 = page1.DrawRectangle(4, -1, 5, 1);

            CustomPropertyHelper.Set(s1, "S1P1", "1");
            CustomPropertyHelper.Set(s2, "S2P1", "2");
            CustomPropertyHelper.Set(s2, "S2P2", "3");
            //set nothing for s3
            CustomPropertyHelper.Set(s4, "S3P1", "4");
            CustomPropertyHelper.Set(s4, "S3P2", "5");
            CustomPropertyHelper.Set(s4, "S3P3", "6");

            var query = new SectionsQuery();

            var prop_sec = query.AddSubQuery(IVisio.VisSectionIndices.visSectionProp);
            var value_col = prop_sec.Columns.Add(VA.ShapeSheet.SrcConstants.CustomPropValue,"Value");

            var shapeids = new[] { s1.ID, s2.ID, s3.ID, s4.ID };

            var data = query.GetFormulasAndResults(page1, shapeids);

            Assert.AreEqual(4, data.Count);
            Assert.AreEqual(1, data[0].Sections[prop_sec].Rows.Count);
            Assert.AreEqual(2, data[1].Sections[prop_sec].Rows.Count);
            Assert.AreEqual(0, data[2].Sections[prop_sec].Rows.Count);
            Assert.AreEqual(3, data[3].Sections[prop_sec].Rows.Count);

            Assert.AreEqual("\"1\"", data[0].Sections[prop_sec].Rows[0].Cells[0].Formula);
            Assert.AreEqual("\"2\"", data[1].Sections[prop_sec].Rows[0].Cells[0].Formula);
            Assert.AreEqual("\"3\"", data[1].Sections[prop_sec].Rows[1].Cells[0].Formula);
            Assert.AreEqual("\"4\"", data[3].Sections[prop_sec].Rows[0].Cells[0].Formula);
            Assert.AreEqual("\"5\"", data[3].Sections[prop_sec].Rows[1].Cells[0].Formula);
            Assert.AreEqual("\"6\"", data[3].Sections[prop_sec].Rows[2].Cells[0].Formula);


            Assert.AreEqual( "1", data[0].Sections[prop_sec].Rows[0].Cells[0].Result);
            Assert.AreEqual( "2", data[1].Sections[prop_sec].Rows[0].Cells[0].Result);
            Assert.AreEqual( "3", data[1].Sections[prop_sec].Rows[1].Cells[0].Result);
            Assert.AreEqual( "4", data[3].Sections[prop_sec].Rows[0].Cells[0].Result);
            Assert.AreEqual( "5", data[3].Sections[prop_sec].Rows[1].Cells[0].Result);
            Assert.AreEqual( "6", data[3].Sections[prop_sec].Rows[2].Cells[0].Result);

            page1.Delete(0);
        }


        [TestMethod]
        public void ShapeSheet_Query_Demo_MultipleShapes()
        {
            var page1 = this.GetNewPage(new VisioAutomation.Geometry.Size(10, 10));

            // draw a simple shape
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(4, 4, 6, 6);
            var s3 = page1.DrawRectangle(5, 5, 7, 7);

            var shapeids = new List<int> { s1.ID, s2.ID, s3.ID };

            Assert.AreEqual(3, page1.Shapes.Count);

            var query = new CellQuery();
            var col_pinx = query.Columns.Add(VA.ShapeSheet.SrcConstants.XFormPinX, "PinX");
            var col_piny = query.Columns.Add(VA.ShapeSheet.SrcConstants.XFormPinY, "PinY");

            var data_formulas = query.GetFormulas(page1, shapeids);
            var data_results = query.GetResults<double>(page1, shapeids);

            var expected_formulas = new[,]
                                      {
                                          {"1 in", "1 in"},
                                          {"5 in", "5 in"},
                                          {"6 in", "6 in"}
                                      };

            var expected_results = new[,]
                                      {
                                          {1.0, 1.0},
                                          {5.0, 5.0},
                                          {6.0, 6.0}
                                      };


            for (int row = 0; row < data_results.Count; row++)
            {
                for (int col = 0; col < query.Columns.Count; col++)
                {
                    Assert.AreEqual(expected_formulas[row,col], data_formulas[row].Cells[col]);
                    Assert.AreEqual(expected_results[row,col], data_results[row].Cells[col]);
                }
            }

            page1.Delete(0);
        }

        [TestMethod]
        public void ShapeSheet_Query_Demo_MultipleShapes_Verify_Out_Of_order()
        {
            var page1 = this.GetNewPage(new VisioAutomation.Geometry.Size(10, 10));

            // draw a simple shape
            var sa = page1.DrawRectangle(-1, -1, 0, 0);
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var sb = page1.DrawRectangle(-1, -1, 0, 0);
            var s2 = page1.DrawRectangle(4, 4, 6, 6);
            var s3 = page1.DrawRectangle(5, 5, 7, 7);

            // notice that the shapes are created as 0, 1,2,3
            // but are queried as 2, 3, 1
            var shapeids = new List<int> { s2.ID, s3.ID, s1.ID };

            Assert.AreEqual(5, page1.Shapes.Count);

            var query = new CellQuery();
            var col_pinx = query.Columns.Add(VA.ShapeSheet.SrcConstants.XFormPinX, "PinX");
            var col_piny = query.Columns.Add(VA.ShapeSheet.SrcConstants.XFormPinY, "PinY");

            var data_formulas = query.GetFormulas(page1, shapeids);
            var data_results = query.GetResults<double>(page1, shapeids);

            var expected_formulas = new[,]
                                      {
                                          {"5 in", "5 in"},
                                          {"6 in", "6 in"},
                                          {"1 in", "1 in"}
                                      };

            var expected_results = new[,]
                                      {
                                          {5.0, 5.0},
                                          {6.0, 6.0},
                                          {1.0, 1.0}
                                      };


            for (int row = 0; row < data_results.Count; row++)
            {
                for (int col = 0; col < query.Columns.Count; col++)
                {
                    Assert.AreEqual(expected_formulas[row, col], data_formulas[row].Cells[col]);
                    Assert.AreEqual(expected_results[row, col], data_results[row].Cells[col]);
                }
            }

            page1.Delete(0);
        }


        [TestMethod]
        public void ShapeSheet_Query_NonExistentSections()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(2, 1, 3, 3);
            var s3 = page1.DrawRectangle(3, 1, 4, 2);
            var s4 = page1.DrawRectangle(4, -1, 5, 1);

            var shapes = new[] {s1, s2, s3, s4};
            var shapeids = shapes.Select(s => s.ID).ToList();

            // First verify that none of the shapes have the controls section locally or otherwise
            foreach (var s in shapes)
            {
                Assert.AreEqual(0, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 1]);
                Assert.AreEqual(0, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 0]);
            }

            // Try to retrieve the control cells rows for each shape, every shape should return zero rows
            foreach (var s in shapes)
            {
                var r1 = ControlCells.GetCells(s);
                Assert.AreEqual(0,r1.Count);
            }

            // Try to retrieve the control cells rows for all shapes at once, every shape should return a collection of zero rows
            var r2 = ControlCells.GetCells(page1, shapeids);
            Assert.AreEqual(shapes.Count(),r2.Count);
            for (int i = 0; i < shapes.Count();i++)
            {
                Assert.AreEqual(0,r2[i].Count);
            }

            // Add a Controls row to shape2
            var cc = new ControlCells();
            ControlHelper.Add(s2, cc);

            // Now verify that none of the shapes *except s2* have the controls section locally or otherwise
            foreach (var s in shapes)
            {
                if (s != s2)
                {
                    Assert.AreEqual(0, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 1]);
                    Assert.AreEqual(0, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 0]);
                }
                else
                {
                    Assert.AreEqual(-1, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 1]);
                    Assert.AreEqual(-1, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 0]);
                }
            }

            // Try to retrieve the control cells rows for each shape, every shape should return zero rows *except for s2*
            foreach (var s in shapes)
            {
                if (s != s2)
                {
                    var r1 = ControlCells.GetCells(s);
                    Assert.AreEqual(0, r1.Count);
                }
                else
                {
                    var r1 = ControlCells.GetCells(s);
                    Assert.AreEqual(1, r1.Count);
                }
            }

            // Try to retrieve the control cells rows for all shapes at once, every shape *except s2* should return a collection of zero rows
            var r3 = ControlCells.GetCells(page1, shapeids);
            Assert.AreEqual(shapes.Count(), r3.Count);
            for (int i = 0; i < shapes.Count(); i++)
            {
                if (shapes[i] != s2)
                {
                    Assert.AreEqual(0, r3[i].Count);
                }
                else
                {
                    Assert.AreEqual(1, r3[i].Count);
                }
            }

            page1.Delete(0);
        }

        public bool section_is_skippable( VA.ShapeSheet.Src src)
        {
            bool can_skip = (src.Section == (short)IVisio.VisSectionIndices.visSectionFirst)
                         || (src.Section == (short)IVisio.VisSectionIndices.visSectionFirstComponent)
                         || (src.Section == (short)IVisio.VisSectionIndices.visSectionLast)
                         || (src.Section == (short)IVisio.VisSectionIndices.visSectionInval)
                         || (src.Section == (short)IVisio.VisSectionIndices.visSectionNone)
                         || (src.Section == (short)IVisio.VisSectionIndices.visSectionFirst)
                         || (src.Section == (short)IVisio.VisSectionIndices.visSectionLastComponent);
            return can_skip;
        }

        public static Dictionary<string, Src> GetSrcDictionary()
        {
            var srcconstants_t = typeof(SrcConstants);

            var binding_flags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Static;

            // find all the static properties that return Src types
            var src_type = typeof(Src);
            var props = srcconstants_t.GetProperties(binding_flags)
                .Where(p => p.PropertyType == src_type);

            var fields_name_to_value = new Dictionary<string, Src>();
            foreach (var propinfo in props)
            {
                var src = (Src)propinfo.GetValue(null, null);
                var name = propinfo.Name;
                fields_name_to_value[name] = src;
            }

            return fields_name_to_value;
        }

        [TestMethod]
        public void ShapeSheet_Query_TestDuplicates()
        {
            // Ensure that duplicate cells are caught
            var q1 = new CellQuery();
            q1.Columns.Add(VA.ShapeSheet.SrcConstants.XFormPinX, "PinX");

            bool caught_exc1 = false;
            try
            {
                q1.Columns.Add(VA.ShapeSheet.SrcConstants.XFormPinX, "PinX");
            }
            catch (System.ArgumentException)
            {
                caught_exc1 = true;
            }

            Assert.IsTrue(caught_exc1);

            // Ensure that duplicate sections are caught

            var q2 = new SectionsQuery();
            q2.AddSubQuery(IVisio.VisSectionIndices.visSectionObject);

            bool caught_exc2 = false;
            try
            {
                q2.AddSubQuery(IVisio.VisSectionIndices.visSectionObject);
            }
            catch (System.ArgumentException)
            {
                caught_exc2 = true;
            }

            Assert.IsTrue(caught_exc2);

            // Ensure that Duplicates in Section Queries Are caught - 
            var q3 = new SectionsQuery();
            var sec = q3.AddSubQuery(IVisio.VisSectionIndices.visSectionObject);
            sec.Columns.Add(VA.ShapeSheet.SrcConstants.XFormPinX,"PinX");
            bool caught_exc3 = false;
            try
            {
                sec.Columns.Add(VA.ShapeSheet.SrcConstants.XFormPinX, "PinX");
            }
            catch (System.ArgumentException)
            {
                caught_exc3 = true;
            }

            Assert.IsTrue(caught_exc3);
        }
    }
}
