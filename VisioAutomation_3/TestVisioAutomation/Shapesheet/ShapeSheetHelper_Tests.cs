using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Diagnostics;

namespace TestVisioAutomation
{
    [TestClass]
    public class ShapeSheetHelperTests_Query : VisioAutomationTest
    {
        private static VA.Metadata.MetadataDB mdx = VA.Metadata.MetadataDB.Load();

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void SpotCheck1()
        {
            var c1 = VA.ShapeSheet.ShapeSheetHelper.TryGetSRCFromName("EndArrow").Value;
            var c2 = VA.ShapeSheet.SRCConstants.EndArrow;

            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(c2, c1);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void CheckAutomationCellNamesAgainstVA()
        {
            var app = GetVisioApplication();
            var documents = app.Documents;
            var doc1 = this.GetNewDoc();
            var page1 = doc1.Pages[1];
            var shape1 = page1.DrawRectangle(0.3, 0, 2.5, 1.7);

            using (var s = app.CreateUndoScope())
            {
                shape1.CellsU["FillForegnd"].FormulaU = "rgb(255,134,78)";
                shape1.CellsU["FillBkgnd"].FormulaU = "rgb(255,134,98)";
                VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(shape1, "custprop1", "value1");
                VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(shape1, "UserDefinedCell1", "Value1", "Prompt1");
                var ctrl = new VA.Controls.ControlCells();
                ctrl.X = "Width*0.5";
                ctrl.Y = "Width*0.75";
                VA.Controls.ControlHelper.AddControl(shape1, ctrl);
                var h1 = shape1.Hyperlinks.Add();
                h1.Address = "http://microsoft/com";

                var t0 = new VA.Text.Markup.TextElement();
                t0.TextFormat.FontSize = VA.Convert.PointsToInches(36);
                var t01 = t0.AppendText("HELLO");
                var t1 = t0.AppendNewElement("W\nO\nR\nL\nD");
                t1.TextFormat.Indent = 1.0;
                t1.TextFormat.FontSize = VA.Convert.PointsToInches(15.0);
                t0.AppendText("FOOBR");

                //t0.SetShapeText(shape1);
                shape1.Text = "0123456789\n0123456789\n0123456789\n0123456789\n01234567890";
                var fmt1 = new VA.Text.CharacterFormatCells();
                fmt1.Transparency = 0.5;
                VA.Text.TextHelper.SetFormat(shape1,fmt1,5,10);

                var fmt2 = new VA.Text.ParagraphFormatCells();
                fmt2.IndentLeft = 1.0;
                VA.Text.TextHelper.SetFormat(shape1,fmt2,1,10);
                VA.Text.TextHelper.SetFormat(shape1, fmt2,20, 30);

                var hlink1=shape1.Hyperlinks.Add();
                var stops = new[]
                                {
                                    new VA.Text.TabStop(0.1, VA.Text.TabStopAlignment.Left),
                                    new VA.Text.TabStop(0.2, VA.Text.TabStopAlignment.Right)
                                };

                VA.Text.TextHelper.SetTabStops(shape1,stops);

            }

            System.Threading.Thread.Sleep(1000);

            var failures = new List<string>();

            foreach (var md_sec in mdx.Sections)
            {
                short sec_index = (short) mdx.GetAutomationConstantByName(md_sec.Enum).GetValueAsInt();
                Debug.WriteLine(md_sec.DisplayName);
                var cells = mdx.Cells.Where(c => c.SectionIndex == md_sec.Enum).Where( c=>c.Object.Contains("shape")).ToList();
                foreach (var cellinfo in EnumCellsInSection(shape1, sec_index))
                {

                    if (cellinfo.NameVisioInterop != cellinfo.NameFromVisioAutomation)
                    {
                        string msg = string.Format(" {0}!={1}  {2}", cellinfo.NameVisioInterop,
                                                   cellinfo.NameFromVisioAutomation, cellinfo.SRC.ToString());
                        failures.Add(msg);
                       
                    }
                }
            }

            if (failures.Count > 0)
            {
                string s = string.Join("\r\n", failures.ToArray());
                //Assert.Fail(s);
            }

            doc1.Close(true);
            app.Quit(true);
        }

        private IEnumerable<CellInfo> EnumCellsInSection(IVisio.Shape shape, short section_index)
        {
            if (0 == shape.SectionExists[section_index, 1])
            {
                yield break;
            }
            var sec = shape.Section[section_index];

            int num_rows = GetCorrectedRowCount(shape, section_index);
            Debug.WriteLine("Num Rows={0}",num_rows);
            for (int r = 0; r < num_rows; r++)
            {
                short row_index = GetCorrectedRowIndex(section_index, r);

                var row = sec[row_index];
                int num_cells = shape.RowsCellCount[section_index,row_index];
                for (int c = 0; c < num_cells; c++)
                {
                    var cell = row[c];
                    var cell_name = cell.Name;
                    var cell_src = new VA.ShapeSheet.SRC(cell.Section, cell.Row, cell.Column);

                    var xcellsrc = VA.ShapeSheet.ShapeSheetHelper.TryGetSRCFromName(cell_name);

                    var ci = new CellInfo();

                    if (!xcellsrc.HasValue)
                    {
                        xcellsrc = new VA.ShapeSheet.SRC(-1, -1, -1);
                        ci.NameFromVisioAutomation = "TBD";
                    }
                    else
                    {
                        ci.NameFromVisioAutomation = VA.ShapeSheet.ShapeSheetHelper.TryGetNameFromSRC(cell_src);
                    }


                    ci.NameVisioInterop = cell_name;
                    ci.SRC = cell_src;

                    ci.XSRC = xcellsrc.Value;

                    ci.Formula = cell.FormulaU;
                    ci.Result = cell.Result[IVisio.tagVisUnitCodes.visNoCast];

                    yield return ci;

                }
            }
        }

        private short GetCorrectedRowIndex(short section_index, int r)
        {
            short row_index = (short)(r + 0);

            if (section_index == (short)IVisio.VisSectionIndices.visSectionObject)
            {
                row_index += 1;
            }
            return row_index;
        }

        private int GetCorrectedRowCount(IVisio.Shape shape, short section_index)
        {
            int num_rows = shape.RowCount[section_index];
            if (section_index == (short)IVisio.VisSectionIndices.visSectionObject)
            {
                if (num_rows < 3)
                {
                    num_rows += 1;                    
                }
            }
            return num_rows;
        }
    }
}
