using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ShapeSheet_Query : VisioAutomationTest
    {
        public static VA.ShapeSheet.SRC cell_fg = VA.ShapeSheet.SRCConstants.FillForegnd;
        public static VA.ShapeSheet.SRC cell_bg = VA.ShapeSheet.SRCConstants.FillBkgnd;
        public static VA.ShapeSheet.SRC cell_pat = VA.ShapeSheet.SRCConstants.FillPattern;


        public static List<T[]> GetRowsInGroup<T>(VA.ShapeSheet.Data.Table<T> table, int group)
        {
            var g = table.Groups[group];
            return g.RowIndices.Select(i => GetRow(table, i)).ToList();
        }

        public static T[] GetRow<T>(VA.ShapeSheet.Data.Table<T> table, int row)
        {
            var a = new T[table.ColumnCount];
            for (int i = 0; i < table.ColumnCount; i++)
            {
                a[i] = table[row, i];
            }
            return a;
        }


        [TestMethod]
        public void Verify_Shape_GetResults_For_Multiple_Types()
        {
            var app = GetVisioApplication();
            var documents = app.Documents;
            var doc1 = this.GetNewDoc();
            var page1 = doc1.Pages[1];
            page1.SetSize(this.StandardPageSize);

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

            var src_fg = VA.ShapeSheet.SRCConstants.FillForegnd;
            var src_bg = VA.ShapeSheet.SRCConstants.FillBkgnd;
            var src_filpat = VA.ShapeSheet.SRCConstants.FillPattern;

            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_fg = query.AddColumn(src_fg);
            var col_bg = query.AddColumn(src_bg);
            var col_filpat = query.AddColumn(src_filpat);

            var shapeids = new[] {s1_id};

            var formulas = query.GetFormulas(page1, shapeids);

            // now verify that the formulas were actually set
            Assert.AreEqual("RGB(255,0,0)", formulas[0,col_fg]);
            Assert.AreEqual("RGB(0,0,255)", formulas[0,col_bg]);
            Assert.AreEqual("40", formulas[0,col_filpat]);

            // now retrieve the results with GetResults as floats
            var float_results = query.GetResults<double>(page1,shapeids);
            Assert.IsNotNull(float_results);
            Assert.AreEqual(24.0, float_results[0, col_fg]);
            Assert.AreEqual(25.0, float_results[0, col_bg]);
            Assert.AreEqual(40.0, float_results[0,col_filpat]);

            // now retrieve the results with GetResults as ints
            var int_results = query.GetResults<int>(page1,shapeids);
            Assert.AreEqual(24, int_results[0, col_fg]);
            Assert.AreEqual(25, int_results[0, col_bg]);
            Assert.AreEqual(40, int_results[0, col_filpat]);

            // now retrieve the results with GetResults as strings

            var string_results = query.GetResults<string>(page1,shapeids);

            Assert.AreEqual("RGB(255, 0, 0)", string_results[0, col_fg]);
            Assert.AreEqual("RGB(0, 0, 255)", string_results[0, col_bg]);
            Assert.AreEqual("40", string_results[0, col_filpat]);

            page1.Delete(0);
            doc1.Close(true);
        }

        [TestMethod]
        public void Verify_Page_GetResults_for_Multiple_Types()
        {
            var page1 = GetNewPage();

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

            var src_fg = VA.ShapeSheet.SRCConstants.FillForegnd;
            var src_bg = VA.ShapeSheet.SRCConstants.FillBkgnd;
            var src_filpat = VA.ShapeSheet.SRCConstants.FillPattern;

            // now retrieve the formulas with GetFormulas

            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_fg = query.AddColumn(src_fg);
            var col_bg = query.AddColumn(src_bg);
            var col_filpat = query.AddColumn(src_filpat);

            var shapeids = new[] {s1_id};

            var formulas = query.GetFormulas(page1, shapeids);

            // now verify that the formulas were actually set
            Assert.AreEqual("2",  formulas[0,col_fg]);
            Assert.AreEqual("3",  formulas[0,col_bg]);
            Assert.AreEqual("40", formulas[0,col_filpat]);

            // now retrieve the results with GetResults as floats

            var float_results = query.GetResults<double>(page1,shapeids);
            Assert.AreEqual(2.0, float_results[0, col_fg]);
            Assert.AreEqual(3.0, float_results[0, col_bg]);
            Assert.AreEqual(40.0, float_results[0, col_filpat]);

            // now retrieve the results with GetResults as ints
            var int_results = query.GetResults<int>(page1,shapeids);

            Assert.AreEqual(2, int_results[0, col_fg]);
            Assert.AreEqual(3, int_results[0, col_bg]);
            Assert.AreEqual(40, int_results[0, col_filpat]);

            // now retrieve the results with GetResults as strings
            var string_results = query.GetResults<string>(page1,shapeids);
            Assert.AreEqual("2", string_results[0, col_fg]);
            Assert.AreEqual("3", string_results[0, col_bg]);
            Assert.AreEqual("40", string_results[0, col_filpat]);

            page1.Delete(0);
        }

        [TestMethod]
        public void Verify_CellQuery_Grouping()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(10, 10));

            // draw a simple shape
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(4, 4, 6, 6);
            var s3 = page1.DrawRectangle(5, 5, 7, 7);

            var shapeids = new List<int> { s1.ID, s2.ID, s3.ID };

            Assert.AreEqual(3, page1.Shapes.Count);

            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_pinx = query.AddColumn(VA.ShapeSheet.SRCConstants.PinX);
            var col_piny = query.AddColumn(VA.ShapeSheet.SRCConstants.PinY);

            var r = query.GetFormulasAndResults<double>(page1, shapeids);

            // Check the grouping
            Assert.AreEqual(shapeids.Count(), r.RowCount); // the total number of rows should match the number of shapeids
            Assert.AreEqual(shapeids.Count(), r.Groups.Count); // the total number of groups should be the number of shapes we asked for

            var expected_pinpos = new List<VA.Drawing.Point>
                                      {
                                          new VA.Drawing.Point(1, 1),
                                          new VA.Drawing.Point(5, 5),
                                          new VA.Drawing.Point(6, 6)
                                      };

            var actual_pinpos = new List<VA.Drawing.Point>(r.RowCount);
            foreach (var row in Enumerable.Range(0, r.RowCount))
            {
                 var p = new VA.Drawing.Point(
                    r[row, col_pinx].Result,
                    r[row, col_piny].Result);
                actual_pinpos.Add(p);
            }

            Assert.AreEqual(expected_pinpos[0], actual_pinpos[0]);
            Assert.AreEqual(expected_pinpos[1], actual_pinpos[1]);
            Assert.AreEqual(expected_pinpos[2], actual_pinpos[2]);
            page1.Delete(0);
        }

        [TestMethod]
        public void Demo_CellQuery_Usage_for_Formulas_and_Results()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(10, 10));

            // draw a simple shape
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(4, 4, 6, 6);
            var s3 = page1.DrawRectangle(5, 5, 7, 7);

            var shapeids = new List<int> { s1.ID, s2.ID, s3.ID };

            Assert.AreEqual(3, page1.Shapes.Count);

            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_pinx = query.AddColumn(VA.ShapeSheet.SRCConstants.PinX);
            var col_piny = query.AddColumn(VA.ShapeSheet.SRCConstants.PinY);

            var r = query.GetFormulasAndResults<double>(page1, shapeids);

            var expected_formulas = new [,]
                                      {
                                          {"1 in", "1 in"},
                                          {"5 in", "5 in"},
                                          {"6 in", "6 in"}
                                      };

            var expected_results = new [,]
                                      {
                                          {1.0, 1.0},
                                          {5.0, 5.0},
                                          {6.0, 6.0}
                                      };


            for (int row = 0; row < r.RowCount; row++)
            {
                for (int col = 0; col < r.ColumnCount; col++)
                {
                    Assert.AreEqual(expected_formulas[row, col], r[row, col].Formula);
                    Assert.AreEqual(expected_results[row, col], r[row, col].Result);
                }
            }

            page1.Delete(0);
        }


        [TestMethod]
        public void Demo_SectionQuery_Grouping()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(2, 1, 3, 3);
            var s3 = page1.DrawRectangle(3, 1, 4, 2);
            var s4 = page1.DrawRectangle(4, -1, 5, 1);

            VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(s1, "S1P1", "1");
            VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(s2, "S2P1", "2");
            VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(s2, "S2P2", "3");
            //set nothing for s3
            VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(s4, "S3P1", "4");
            VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(s4, "S3P2", "5");
            VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(s4, "S3P3", "6");

            var query = new VA.ShapeSheet.Query.SectionQuery((short)IVisio.VisSectionIndices.visSectionProp);
            query.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Value.Cell);

            var shapeids = new[] { s1.ID, s2.ID, s3.ID, s4.ID };

            var table = query.GetFormulasAndResults<double>(
                page1,
                shapeids);

            Assert.AreEqual(4, table.Groups.Count);
            Assert.AreEqual(1, table.Groups[0].Count);
            Assert.AreEqual(2, table.Groups[1].Count);
            Assert.AreEqual(0, table.Groups[2].Count);
            Assert.AreEqual(3, table.Groups[3].Count);

            var gf0 = GetRowsInGroup(table, 0);
            var gf1 = GetRowsInGroup(table, 1);
            var gf2 = GetRowsInGroup(table, 2);
            var gf3 = GetRowsInGroup(table, 3);


            Assert.AreEqual("\"1\"", gf0[0][0].Formula);
            Assert.AreEqual("\"2\"", gf1[0][0].Formula);
            Assert.AreEqual("\"3\"", gf1[1][0].Formula);
            Assert.AreEqual("\"4\"", gf3[0][0].Formula);
            Assert.AreEqual("\"5\"", gf3[1][0].Formula);
            Assert.AreEqual("\"6\"", gf3[2][0].Formula);


            Assert.AreEqual(1.0, gf0[0][0].Result);
            Assert.AreEqual(2.0, gf1[0][0].Result);
            Assert.AreEqual(3.0, gf1[1][0].Result);
            Assert.AreEqual(4.0, gf3[0][0].Result);
            Assert.AreEqual(5.0, gf3[1][0].Result);
            Assert.AreEqual(6.0, gf3[2][0].Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void Verify_SectionQuery_Grouping()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(2, 1, 3, 3);
            var s3 = page1.DrawRectangle(3, 1, 4, 2);
            var s4 = page1.DrawRectangle(4, -1, 5, 1);

            var query = new VA.ShapeSheet.Query.SectionQuery((short)IVisio.VisSectionIndices.visSectionProp);
            query.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Value.Cell);
            var shapeids = new[] { s1.ID, s2.ID, s3.ID, s4.ID };

            var table = query.GetFormulasAndResults<double>(
                page1, shapeids);
            
            Assert.AreEqual(4, table.Groups.Count);
            Assert.AreEqual(0, table.Groups[0].Count);
            Assert.AreEqual(0, table.Groups[1].Count);
            Assert.AreEqual(0, table.Groups[2].Count);
            Assert.AreEqual(0, table.Groups[3].Count);

            page1.Delete(0);
        }

        private static VA.ShapeSheet.Query.CellQuery BuildCellQuery(IList<VA.ShapeSheet.SRC> srcs)
        {
            var query = new VA.ShapeSheet.Query.CellQuery();
            foreach (var src in srcs)
            {
                query.AddColumn(src);
            }
            return query;
        }

        [TestMethod]
        public void Verify_CellQuery_Results_for_multiple_types()
        {
            var page1 = GetNewPage();

            // draw a simple shape
            var s1 = page1.DrawRectangle(this.StandardPageSizeRect);

            // format it with setformulas
            var fg_cell = s1.Cells["FillForegnd"];
            var bg_cell = s1.Cells["FillBkgnd"];
            var pat_cell = s1.Cells["FillPattern"];

            fg_cell.ResultIU = 2.0; //red
            bg_cell.ResultIU = 3.0; //green
            pat_cell.ResultIU = 40.0;

            // now retrieve the formulas with GetFormulas


            var query = BuildCellQuery(new[] { cell_fg, cell_bg, cell_pat });
            var formulas = query.GetFormulas(s1);

            // now verify that the formulas were actually set
            Assert.AreEqual("2", formulas[0, 0]);
            Assert.AreEqual("3", formulas[0, 1]);
            Assert.AreEqual("40", formulas[0, 2]);

            // now retrieve the results with GetResults as floats

            var float_results = query.GetResults<double>(s1);
            Assert.AreEqual(2.0, float_results[0, 0]);
            Assert.AreEqual(3.0, float_results[0, 1]);
            Assert.AreEqual(40.0, float_results[0, 2]);

            // now retrieve the results with GetResults as ints
            var int_results = query.GetResults<int>(s1);
            Assert.AreEqual(2, int_results[0, 0]);
            Assert.AreEqual(3, int_results[0, 1]);
            Assert.AreEqual(40, int_results[0, 2]);

            // now retrieve the results with GetResults as strings
            var string_results = query.GetResults<string>(s1);
            Assert.AreEqual("2", string_results[0, 0]);
            Assert.AreEqual("3", string_results[0, 1]);
            Assert.AreEqual("40", string_results[0, 2]);

            page1.Delete(0);
        }

        [TestMethod]
        public void Verify_SectionQuery_With_NonExistentSections()
        {
            var page1 = GetNewPage();
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
                var r1 = VA.Controls.ControlCells.GetCells(s);
                Assert.AreEqual(0,r1.Count);
            }

            // Try to retrieve the control cells rows for all shapes at once, every shape should return a collection of zero rows
            var r2 = VA.Controls.ControlCells.GetCells(page1, shapeids);
            Assert.AreEqual(shapes.Count(),r2.Count);
            for (int i = 0; i < shapes.Count();i++)
            {
                Assert.AreEqual(0,r2[i].Count);
            }

            // Add a Controls row to shape2
            var cc = new VA.Controls.ControlCells();
            VA.Controls.ControlHelper.AddControl(s2, cc);

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
                    var r1 = VA.Controls.ControlCells.GetCells(s);
                    Assert.AreEqual(0, r1.Count);
                }
                else
                {
                    var r1 = VA.Controls.ControlCells.GetCells(s);
                    Assert.AreEqual(1, r1.Count);
                }
            }

            // Try to retrieve the control cells rows for all shapes at once, every shape *except s2* should return a collection of zero rows
            var r3 = VA.Controls.ControlCells.GetCells(page1, shapeids);
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

    }
}
