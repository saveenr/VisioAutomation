using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VTest.Core.ShapeSheet
{
    [TestClass]
    public class ShapeSheetQueryTests : VisioAutomationTest
    {
        [TestMethod]
        public void ShapeSheet_Query_GetResults_SingleShape()
        {
            var doc1 = this.GetNewDoc();
            var page1 = doc1.Pages[1];
            SetPageSize(page1, this.StandardPageSize);

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

            var query = new VASS.Query.CellQuery();
            var col_fg = query.Columns.Add(VisioAutomation.Core.SrcConstants.FillForeground);
            var col_bg = query.Columns.Add(VisioAutomation.Core.SrcConstants.FillBackground);
            var col_filpat = query.Columns.Add(VisioAutomation.Core.SrcConstants.FillPattern);

            var shapeids = new[] {s1_id};

            var formulas = query.GetFormulas(page1, shapeids);

            // now verify that the formulas were actually set
            Assert.AreEqual("RGB(255,0,0)", formulas[0][col_fg]);
            Assert.AreEqual("RGB(0,0,255)", formulas[0][col_bg]);
            Assert.AreEqual("40", formulas[0][col_filpat]);

            // now retrieve the results with GetResults as floats
            var float_results = query.GetResults<double>(page1,shapeids);
            Assert.IsNotNull(float_results);
            Assert.AreEqual(40.0, float_results[0][col_filpat]);

            // now retrieve the results with GetResults as ints
            var int_results = query.GetResults<int>(page1, shapeids);
            Assert.AreEqual(40, int_results[0][col_filpat]);

            // now retrieve the results with GetResults as strings

            var string_results = query.GetResults<string>(page1, shapeids);

            Assert.AreEqual("RGB(255, 0, 0)", string_results[0][col_fg]);
            Assert.AreEqual("RGB(0, 0, 255)", string_results[0][col_bg]);
            Assert.AreEqual("40", string_results[0][col_filpat]);

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

            // now retrieve the formulas with GetFormulas

            var query = new VASS.Query.CellQuery();
            var col_fg = query.Columns.Add(VisioAutomation.Core.SrcConstants.FillForeground );
            var col_bg = query.Columns.Add(VisioAutomation.Core.SrcConstants.FillBackground );
            var col_filpat = query.Columns.Add(VisioAutomation.Core.SrcConstants.FillPattern);

            var shapeids = new[] {s1_id};

            var formulas = query.GetFormulas(page1, shapeids);

            // now verify that the formulas were actually set
            Assert.AreEqual("2", formulas[0][col_fg]);
            Assert.AreEqual("3", formulas[0][col_bg]);
            Assert.AreEqual("40", formulas[0][col_filpat]);

            // now retrieve the results with GetResults as floats
            var float_results = query.GetResults<double>(page1, shapeids);
            Assert.AreEqual(2.0, float_results[0][col_fg]);
            Assert.AreEqual(3.0, float_results[0][col_bg]);
            Assert.AreEqual(40.0, float_results[0][col_filpat]);

            // now retrieve the results with GetResults as ints
            var int_results = query.GetResults<int>(page1, shapeids);

            Assert.AreEqual(2, int_results[0][col_fg]);
            Assert.AreEqual(3, int_results[0][col_bg]);
            Assert.AreEqual(40, int_results[0][col_filpat]);

            // now retrieve the results with GetResults as strings
            var string_results = query.GetResults<string>(page1,shapeids);
            Assert.AreEqual("2", string_results[0][col_fg]);
            Assert.AreEqual("3", string_results[0][col_bg]);
            Assert.AreEqual("40", string_results[0][col_filpat]);

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

            int cp_type = 0; // 0 for string

            VA.Shapes.CustomPropertyHelper.Set(s1, "S1P1", "\"1\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s2, "S2P1", "\"2\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s2, "S2P2", "\"3\"", cp_type);
            //set nothing for s3
            VA.Shapes.CustomPropertyHelper.Set(s4, "S3P1", "\"4\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s4, "S3P2", "\"5\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s4, "S3P3", "\"6\"", cp_type);

            var sec_query = new VASS.Query.SectionQuery();

            var sec_cols = sec_query.Add(IVisio.VisSectionIndices.visSectionProp);
            var value_col = sec_cols.Add(VisioAutomation.Core.SrcConstants.CustomPropValue);

            var shapeidpairs = VA.Core.ShapeIDPairs.FromShapes( s1, s2, s3, s4 );

            var data = sec_query.GetFormulas(page1, shapeidpairs);
            var data2 = sec_query.GetResults<string>(page1, shapeidpairs);

            int shape0_index = 0;
            int shape1_index = 1;
            int shape2_index = 2;
            int shape3_index = 3;
            int section0_index = 0;

            Assert.AreEqual(4, data.Count);
            Assert.AreEqual(1, data[shape0_index][section0_index].Count);
            Assert.AreEqual(2, data[shape1_index][section0_index].Count);
            Assert.AreEqual(0, data[shape2_index][section0_index].Count);
            Assert.AreEqual(3, data[3][0].Count);

            Assert.AreEqual("\"1\"", data[shape0_index][section0_index][0][0]);
            Assert.AreEqual("\"2\"", data[shape1_index][section0_index][0][0]);
            Assert.AreEqual("\"3\"", data[shape1_index][section0_index][1][0]);
            Assert.AreEqual("\"4\"", data[shape3_index][section0_index][0][0]);
            Assert.AreEqual("\"5\"", data[shape3_index][section0_index][1][0]);
            Assert.AreEqual("\"6\"", data[shape3_index][section0_index][2][0]);


            Assert.AreEqual( "1", data2[shape0_index][section0_index][0][0]);
            Assert.AreEqual( "2", data2[shape1_index][section0_index][0][0]);
            Assert.AreEqual( "3", data2[shape1_index][section0_index][1][0]);
            Assert.AreEqual( "4", data2[shape3_index][section0_index][0][0]);
            Assert.AreEqual( "5", data2[shape3_index][section0_index][1][0]);
            Assert.AreEqual( "6", data2[shape3_index][section0_index][2][0]);

            page1.Delete(0);
        }


        [TestMethod]
        public void ShapeSheet_Query_Demo_MultipleShapes()
        {
            var page1 = this.GetNewPage(new VA.Core.Size(10, 10));

            // draw a simple shape
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(4, 4, 6, 6);
            var s3 = page1.DrawRectangle(5, 5, 7, 7);

            var shapeids = new List<int> { s1.ID, s2.ID, s3.ID };

            Assert.AreEqual(3, page1.Shapes.Count);

            var query = new VASS.Query.CellQuery();
            var col_pinx = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormPinX);
            var col_piny = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormPinY);

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
                    Assert.AreEqual(expected_formulas[row,col], data_formulas[row][col]);
                    Assert.AreEqual(expected_results[row,col], data_results[row][col]);
                }
            }

            page1.Delete(0);
        }

        [TestMethod]
        public void ShapeSheet_Query_Demo_MultipleShapes_Verify_Out_Of_order()
        {
            var page1 = this.GetNewPage(new VA.Core.Size(10, 10));

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

            var query = new VASS.Query.CellQuery();
            var col_pinx = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormPinX);
            var col_piny = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormPinY);

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
                    Assert.AreEqual(expected_formulas[row, col], data_formulas[row][col]);
                    Assert.AreEqual(expected_results[row, col], data_results[row][col]);
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
            var shapeidpairs = VA.Core.ShapeIDPairs.FromShapes(shapes);


            // First verify that none of the shapes have the controls section locally or otherwise
            foreach (var s in shapes)
            {
                Assert.AreEqual(0, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 1]);
                Assert.AreEqual(0, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 0]);
            }

            // Try to retrieve the control cells rows for each shape, every shape should return zero rows
            foreach (var s in shapes)
            {
                var r1 = VA.Shapes.ControlCells.GetCells(s, VisioAutomation.Core.CellValueType.Formula);
                Assert.AreEqual(0,r1.Count);
            }

            // Try to retrieve the control cells rows for all shapes at once, every shape should return a collection of zero rows
            var r2 = VA.Shapes.ControlCells.GetCells(page1, shapeidpairs, VisioAutomation.Core.CellValueType.Formula);
            Assert.AreEqual(shapes.Length,r2.Count);
            for (int i = 0; i < shapes.Length;i++)
            {
                Assert.AreEqual(0,r2[i].Count);
            }

            // Add a Controls row to shape2
            var cc = new VA.Shapes.ControlCells();
            VA.Shapes.ControlHelper.Add(s2, cc);

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
                    var r1 = VA.Shapes.ControlCells.GetCells(s, VisioAutomation.Core.CellValueType.Formula);
                    Assert.AreEqual(0, r1.Count);
                }
                else
                {
                    var r1 = VA.Shapes.ControlCells.GetCells(s, VisioAutomation.Core.CellValueType.Formula);
                    Assert.AreEqual(1, r1.Count);
                }
            }

            // Try to retrieve the control cells rows for all shapes at once, every shape *except s2* should return a collection of zero rows
            var r3 = VA.Shapes.ControlCells.GetCells(page1, shapeidpairs, VisioAutomation.Core.CellValueType.Formula);
            Assert.AreEqual(shapes.Length, r3.Count);
            for (int i = 0; i < shapes.Length; i++)
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

        public bool section_is_skippable(VisioAutomation.Core.Src src)
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

        public static Dictionary<string, VisioAutomation.Core.Src> GetSrcDictionary()
        {
            var srcconstants_t = typeof(VisioAutomation.Core.SrcConstants);

            var binding_flags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Static;

            // find all the static properties that return Src types
            var src_type = typeof(VisioAutomation.Core.Src);
            var props = srcconstants_t.GetProperties(binding_flags)
                .Where(p => p.PropertyType == src_type);

            var fields_name_to_value = new Dictionary<string, VisioAutomation.Core.Src>();
            foreach (var propinfo in props)
            {
                var src = (VisioAutomation.Core.Src)propinfo.GetValue(null, null);
                var name = propinfo.Name;
                fields_name_to_value[name] = src;
            }

            return fields_name_to_value;
        }

        [TestMethod]
        public void ShapeSheet_Query_TestDuplicates()
        {
            // Ensure that duplicate cells are caught
            var cell_query_1 = new VASS.Query.CellQuery();
            cell_query_1.Columns.Add(VisioAutomation.Core.SrcConstants.XFormPinX);

            bool caught_exc1 = false;
            try
            {
                cell_query_1.Columns.Add(VisioAutomation.Core.SrcConstants.XFormPinX);
            }
            catch (System.ArgumentException)
            {
                caught_exc1 = true;
            }

            Assert.IsTrue(caught_exc1);

            // Ensure that duplicate sections are caught

            var sec_query_2 = new VASS.Query.SectionQuery();
            sec_query_2.Add(IVisio.VisSectionIndices.visSectionObject);

            bool caught_exc2 = false;
            try
            {
                sec_query_2.Add(IVisio.VisSectionIndices.visSectionObject);
            }
            catch (System.ArgumentException)
            {
                caught_exc2 = true;
            }

            Assert.IsTrue(caught_exc2);

            // Ensure that Duplicates in Section Queries Are caught - 
            var sec_query_3 = new VASS.Query.SectionQuery();
            var sec_cols = sec_query_3.Add(IVisio.VisSectionIndices.visSectionObject);
            sec_cols.Add(VisioAutomation.Core.SrcConstants.XFormPinX);
            bool caught_exc3 = false;
            try
            {
                sec_cols.Add(VisioAutomation.Core.SrcConstants.XFormPinX);
            }
            catch (System.ArgumentException)
            {
                caught_exc3 = true;
            }

            Assert.IsTrue(caught_exc3);
        }
    }
}
