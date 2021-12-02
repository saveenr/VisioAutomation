using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VASS = VisioAutomation.ShapeSheet;

namespace VTest.Core.ShapeSheet
{
    [MUT.TestClass]
    public class ShapeSheetQueryTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void ShapeSheet_Query_GetResults_SingleShape()
        {
            var doc1 = this.GetNewDoc();
            var page1 = doc1.Pages[1];
            SetPageSize(page1, this.StandardPageSize);

            // draw a simple shape
            var shape_a = page1.DrawRectangle(this.StandardPageSizeRect);
            int shape_a_id = shape_a.ID;

            // format it with setformulas
            var fg_cell = shape_a.Cells["FillForegnd"];
            var bg_cell = shape_a.Cells["FillBkgnd"];
            var pat_cell = shape_a.Cells["FillPattern"];

            fg_cell.FormulaU = "RGB(255,0,0)";
            bg_cell.FormulaU = "RGB(0,0,255)";
            pat_cell.FormulaU = "40";

            // now retrieve the formulas with GetFormulas

            var query = new VASS.Query.CellQuery();
            var col_fg = query.Columns.Add(VA.Core.SrcConstants.FillForeground);
            var col_bg = query.Columns.Add(VA.Core.SrcConstants.FillBackground);
            var col_filpat = query.Columns.Add(VA.Core.SrcConstants.FillPattern);

            var shapeids = new[] {shape_a_id};

            var formulas = query.GetFormulas(page1, shapeids);

            // now verify that the formulas were actually set
            MUT.Assert.AreEqual("RGB(255,0,0)", formulas[0][col_fg]);
            MUT.Assert.AreEqual("RGB(0,0,255)", formulas[0][col_bg]);
            MUT.Assert.AreEqual("40", formulas[0][col_filpat]);

            // now retrieve the results with GetResults as floats
            var float_results = query.GetResults<double>(page1,shapeids);
            MUT.Assert.IsNotNull(float_results);
            MUT.Assert.AreEqual(40.0, float_results[0][col_filpat]);

            // now retrieve the results with GetResults as ints
            var int_results = query.GetResults<int>(page1, shapeids);
            MUT.Assert.AreEqual(40, int_results[0][col_filpat]);

            // now retrieve the results with GetResults as strings

            var string_results = query.GetResults<string>(page1, shapeids);

            MUT.Assert.AreEqual("RGB(255, 0, 0)", string_results[0][col_fg]);
            MUT.Assert.AreEqual("RGB(0, 0, 255)", string_results[0][col_bg]);
            MUT.Assert.AreEqual("40", string_results[0][col_filpat]);

            page1.Delete(0);
            doc1.Close(true);
        }

        [MUT.TestMethod]
        public void ShapeSheet_Query_GetResults_MultipleShapes()
        {
            var page1 = this.GetNewPage();

            // draw a simple shape
            var shape_a = page1.DrawRectangle(this.StandardPageSizeRect);
            int shape_a_id = shape_a.ID;

            // format it with setformulas
            var fg_cell = shape_a.Cells["FillForegnd"];
            var bg_cell = shape_a.Cells["FillBkgnd"];
            var pat_cell = shape_a.Cells["FillPattern"];

            fg_cell.ResultIU = 2.0; //red
            bg_cell.ResultIU = 3.0; //green
            pat_cell.ResultIU = 40.0;

            // now retrieve the formulas with GetFormulas

            var query = new VASS.Query.CellQuery();
            var col_fg = query.Columns.Add(VA.Core.SrcConstants.FillForeground );
            var col_bg = query.Columns.Add(VA.Core.SrcConstants.FillBackground );
            var col_filpat = query.Columns.Add(VA.Core.SrcConstants.FillPattern);

            var shapeids = new[] {shape_a_id};

            var formulas = query.GetFormulas(page1, shapeids);

            // now verify that the formulas were actually set
            MUT.Assert.AreEqual("2", formulas[0][col_fg]);
            MUT.Assert.AreEqual("3", formulas[0][col_bg]);
            MUT.Assert.AreEqual("40", formulas[0][col_filpat]);

            // now retrieve the results with GetResults as floats
            var float_results = query.GetResults<double>(page1, shapeids);
            MUT.Assert.AreEqual(2.0, float_results[0][col_fg]);
            MUT.Assert.AreEqual(3.0, float_results[0][col_bg]);
            MUT.Assert.AreEqual(40.0, float_results[0][col_filpat]);

            // now retrieve the results with GetResults as ints
            var int_results = query.GetResults<int>(page1, shapeids);

            MUT.Assert.AreEqual(2, int_results[0][col_fg]);
            MUT.Assert.AreEqual(3, int_results[0][col_bg]);
            MUT.Assert.AreEqual(40, int_results[0][col_filpat]);

            // now retrieve the results with GetResults as strings
            var string_results = query.GetResults<string>(page1,shapeids);
            MUT.Assert.AreEqual("2", string_results[0][col_fg]);
            MUT.Assert.AreEqual("3", string_results[0][col_bg]);
            MUT.Assert.AreEqual("40", string_results[0][col_filpat]);

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void ShapeSheet_Query_SectionRowHandling()
        {
            var page1 = this.GetNewPage();
            var shape_a = page1.DrawRectangle(0, 0, 2, 2);
            var shape_b = page1.DrawRectangle(2, 1, 3, 3);
            var shape_c = page1.DrawRectangle(3, 1, 4, 2);
            var shape_d = page1.DrawRectangle(4, -1, 5, 1);

            int cp_type = 0; // 0 for string

            VA.Shapes.CustomPropertyHelper.Set(shape_a, "S1P1", "\"1\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(shape_b, "S2P1", "\"2\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(shape_b, "S2P2", "\"3\"", cp_type);
            //set nothing for s3
            VA.Shapes.CustomPropertyHelper.Set(shape_d, "S3P1", "\"4\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(shape_d, "S3P2", "\"5\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(shape_d, "S3P3", "\"6\"", cp_type);

            var sec_query = new VASS.Query.SectionQuery();

            var sec_cols = sec_query.Add(IVisio.VisSectionIndices.visSectionProp);
            var value_col = sec_cols.Add(VA.Core.SrcConstants.CustomPropValue);

            var shapeidpairs = VA.Core.ShapeIDPairs.FromShapes( shape_a, shape_b, shape_c, shape_d );

            var output_formulas = sec_query.GetFormulas(page1, shapeidpairs);
            var output_results_str = sec_query.GetResults<string>(page1, shapeidpairs);

            int shape_a_index = 0;
            int shape_b_index = 1;
            int shape_c_index = 2;
            int shape_d_index = 3;
            int section0_index = 0;

            MUT.Assert.AreEqual(4, output_formulas.Count);
            MUT.Assert.AreEqual(1, output_formulas[shape_a_index][section0_index].Count);
            MUT.Assert.AreEqual(2, output_formulas[shape_b_index][section0_index].Count);
            MUT.Assert.AreEqual(0, output_formulas[shape_c_index][section0_index].Count);
            MUT.Assert.AreEqual(3, output_formulas[3][0].Count);

            MUT.Assert.AreEqual("\"1\"", output_formulas[shape_a_index][section0_index][0][0]);
            MUT.Assert.AreEqual("\"2\"", output_formulas[shape_b_index][section0_index][0][0]);
            MUT.Assert.AreEqual("\"3\"", output_formulas[shape_b_index][section0_index][1][0]);
            MUT.Assert.AreEqual("\"4\"", output_formulas[shape_d_index][section0_index][0][0]);
            MUT.Assert.AreEqual("\"5\"", output_formulas[shape_d_index][section0_index][1][0]);
            MUT.Assert.AreEqual("\"6\"", output_formulas[shape_d_index][section0_index][2][0]);


            MUT.Assert.AreEqual( "1", output_results_str[shape_a_index][section0_index][0][0]);
            MUT.Assert.AreEqual( "2", output_results_str[shape_b_index][section0_index][0][0]);
            MUT.Assert.AreEqual( "3", output_results_str[shape_b_index][section0_index][1][0]);
            MUT.Assert.AreEqual( "4", output_results_str[shape_d_index][section0_index][0][0]);
            MUT.Assert.AreEqual( "5", output_results_str[shape_d_index][section0_index][1][0]);
            MUT.Assert.AreEqual( "6", output_results_str[shape_d_index][section0_index][2][0]);

            page1.Delete(0);
        }


        [MUT.TestMethod]
        public void ShapeSheet_Query_Demo_MultipleShapes()
        {
            var page1 = this.GetNewPage(new VA.Core.Size(10, 10));

            // draw a simple shape
            var shape_a = page1.DrawRectangle(0, 0, 2, 2);
            var shape_b = page1.DrawRectangle(4, 4, 6, 6);
            var shape_c = page1.DrawRectangle(5, 5, 7, 7);

            var shapeids = new List<int> { shape_a.ID, shape_b.ID, shape_c.ID };

            MUT.Assert.AreEqual(3, page1.Shapes.Count);

            var query = new VASS.Query.CellQuery();
            var col_pinx = query.Columns.Add(VA.Core.SrcConstants.XFormPinX);
            var col_piny = query.Columns.Add(VA.Core.SrcConstants.XFormPinY);

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
                    MUT.Assert.AreEqual(expected_formulas[row,col], data_formulas[row][col]);
                    MUT.Assert.AreEqual(expected_results[row,col], data_results[row][col]);
                }
            }

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void ShapeSheet_Query_Demo_MultipleShapes_Verify_Out_Of_order()
        {
            var page1 = this.GetNewPage(new VA.Core.Size(10, 10));

            // draw a simple shape
            var shape_a = page1.DrawRectangle(-1, -1, 0, 0);
            var shape_b = page1.DrawRectangle(0, 0, 2, 2);
            var shape_c = page1.DrawRectangle(-1, -1, 0, 0);
            var shape_d = page1.DrawRectangle(4, 4, 6, 6);
            var shape_e = page1.DrawRectangle(5, 5, 7, 7);

            // notice that the shapes are created as 0, 1,2,3
            // but are queried as 2, 3, 1
            var shapeids = new List<int> { shape_d.ID, shape_e.ID, shape_b.ID };

            MUT.Assert.AreEqual(5, page1.Shapes.Count);

            var query = new VASS.Query.CellQuery();
            var col_pinx = query.Columns.Add(VA.Core.SrcConstants.XFormPinX);
            var col_piny = query.Columns.Add(VA.Core.SrcConstants.XFormPinY);

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
                    MUT.Assert.AreEqual(expected_formulas[row, col], data_formulas[row][col]);
                    MUT.Assert.AreEqual(expected_results[row, col], data_results[row][col]);
                }
            }

            page1.Delete(0);
        }


        [MUT.TestMethod]
        public void ShapeSheet_Query_NonExistentSections()
        {
            var page1 = this.GetNewPage();
            var shape_a = page1.DrawRectangle(0, 0, 2, 2);
            var shape_b = page1.DrawRectangle(2, 1, 3, 3);
            var shape_c = page1.DrawRectangle(3, 1, 4, 2);
            var shape_d = page1.DrawRectangle(4, -1, 5, 1);

            var shapes = new[] {shape_a, shape_b, shape_c, shape_d};
            var shapeidpairs = VA.Core.ShapeIDPairs.FromShapes(shapes);


            // First verify that none of the shapes have the controls section locally or otherwise
            foreach (var s in shapes)
            {
                MUT.Assert.AreEqual(0, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 1]);
                MUT.Assert.AreEqual(0, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 0]);
            }

            // Try to retrieve the control cells rows for each shape, every shape should return zero rows
            foreach (var s in shapes)
            {
                var r1 = VA.Shapes.ControlCells.GetCells(s, VA.Core.CellValueType.Formula);
                MUT.Assert.AreEqual(0,r1.Count);
            }

            // Try to retrieve the control cells rows for all shapes at once, every shape should return a collection of zero rows
            var r2 = VA.Shapes.ControlCells.GetCells(page1, shapeidpairs, VA.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(shapes.Length,r2.Count);
            for (int i = 0; i < shapes.Length;i++)
            {
                MUT.Assert.AreEqual(0,r2[i].Count);
            }

            // Add a Controls row to shape2
            var cc = new VA.Shapes.ControlCells();
            VA.Shapes.ControlHelper.Add(shape_b, cc);

            // Now verify that none of the shapes *except s2* have the controls section locally or otherwise
            foreach (var s in shapes)
            {
                if (s != shape_b)
                {
                    MUT.Assert.AreEqual(0, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 1]);
                    MUT.Assert.AreEqual(0, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 0]);
                }
                else
                {
                    MUT.Assert.AreEqual(-1, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 1]);
                    MUT.Assert.AreEqual(-1, s.SectionExists[(short)IVisio.VisSectionIndices.visSectionControls, 0]);
                }
            }

            // Try to retrieve the control cells rows for each shape, every shape should return zero rows *except for s2*
            foreach (var s in shapes)
            {
                if (s != shape_b)
                {
                    var r1 = VA.Shapes.ControlCells.GetCells(s, VA.Core.CellValueType.Formula);
                    MUT.Assert.AreEqual(0, r1.Count);
                }
                else
                {
                    var r1 = VA.Shapes.ControlCells.GetCells(s, VA.Core.CellValueType.Formula);
                    MUT.Assert.AreEqual(1, r1.Count);
                }
            }

            // Try to retrieve the control cells rows for all shapes at once, every shape *except s2* should return a collection of zero rows
            var r3 = VA.Shapes.ControlCells.GetCells(page1, shapeidpairs, VA.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(shapes.Length, r3.Count);
            for (int i = 0; i < shapes.Length; i++)
            {
                if (shapes[i] != shape_b)
                {
                    MUT.Assert.AreEqual(0, r3[i].Count);
                }
                else
                {
                    MUT.Assert.AreEqual(1, r3[i].Count);
                }
            }

            page1.Delete(0);
        }

        public bool section_is_skippable(VA.Core.Src src)
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

        public static Dictionary<string, VA.Core.Src> GetSrcDictionary()
        {
            var srcconstants_t = typeof(VA.Core.SrcConstants);

            var binding_flags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Static;

            // find all the static properties that return Src types
            var src_type = typeof(VA.Core.Src);
            var props = srcconstants_t.GetProperties(binding_flags)
                .Where(p => p.PropertyType == src_type);

            var fields_name_to_value = new Dictionary<string, VA.Core.Src>();
            foreach (var propinfo in props)
            {
                var src = (VA.Core.Src)propinfo.GetValue(null, null);
                var name = propinfo.Name;
                fields_name_to_value[name] = src;
            }

            return fields_name_to_value;
        }

        [MUT.TestMethod]
        public void ShapeSheet_Query_TestDuplicates()
        {
            // Ensure that duplicate cells cannot be added to a cell query
            var cell_query_1 = new VASS.Query.CellQuery();
            cell_query_1.Columns.Add(VA.Core.SrcConstants.XFormPinX);

            bool caught_exc1 = false;
            try
            {
                cell_query_1.Columns.Add(VA.Core.SrcConstants.XFormPinX);
            }
            catch (System.ArgumentException)
            {
                caught_exc1 = true;
            }

            MUT.Assert.IsTrue(caught_exc1);

            // Ensure that duplicate sections cannot be added to a section query
            
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

            MUT.Assert.IsTrue(caught_exc2);

            // Ensure that Duplicates in Section Queries Are caught - 
            var sec_query_3 = new VASS.Query.SectionQuery();
            var sec_cols = sec_query_3.Add(IVisio.VisSectionIndices.visSectionObject);
            sec_cols.Add(VA.Core.SrcConstants.XFormPinX);
            bool caught_exc3 = false;
            try
            {
                sec_cols.Add(VA.Core.SrcConstants.XFormPinX);
            }
            catch (System.ArgumentException)
            {
                caught_exc3 = true;
            }

            MUT.Assert.IsTrue(caught_exc3);
        }
    }
}
