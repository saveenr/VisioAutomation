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
        public class CellInfo
        {
            public string RealName;
            public VA.ShapeSheet.SRC SRC;
            public string XName;
            public VA.ShapeSheet.SRC XSRC;
            public string Formula;
            public double Result;

        }
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void SpotCheck1()
        {
            var c1 = VA.ShapeSheet.ShapeSheetHelper.TryGetSRCFromName("EndArrow").Value;
            var c2 = VA.ShapeSheet.SRCConstants.EndArrow;

            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(c2, c1);
        }

        static short[] common_section_indices = new[] 
            {     
                (short) IVisio.VisSectionIndices.visSectionAction, 
                (short) IVisio.VisSectionIndices.visSectionAnnotation,
                (short) IVisio.VisSectionIndices.visSectionCharacter,
                (short) IVisio.VisSectionIndices.visSectionConnectionPts,
                (short) IVisio.VisSectionIndices.visSectionControls,
                (short) IVisio.VisSectionIndices.visSectionExport, 
                (short) IVisio.VisSectionIndices.visSectionHyperlink,
                (short) IVisio.VisSectionIndices.visSectionLayer, 
                (short) IVisio.VisSectionIndices.visSectionParagraph,
                (short) IVisio.VisSectionIndices.visSectionProp, 
                (short) IVisio.VisSectionIndices.visSectionReviewer,
                (short) IVisio.VisSectionIndices.visSectionScratch, 
                (short) IVisio.VisSectionIndices.visSectionSmartTag,
                (short) IVisio.VisSectionIndices.visSectionTab, 
                (short) IVisio.VisSectionIndices.visSectionTextField,
                (short) IVisio.VisSectionIndices.visSectionUser, 
                (short) IVisio.VisSectionIndices.visSectionObject  
            };

        static Dictionary<short, string> section_to_name = new Dictionary<short, string>
            {
                { (short) IVisio.VisSectionIndices.visSectionAction, "Action" },
                { (short) IVisio.VisSectionIndices.visSectionAnnotation, "Annotation" },
                { (short) IVisio.VisSectionIndices.visSectionCharacter, "Character" },
                { (short) IVisio.VisSectionIndices.visSectionConnectionPts, "ConnectionPts" },
                { (short) IVisio.VisSectionIndices.visSectionControls, "Controls" },
                { (short) IVisio.VisSectionIndices.visSectionHyperlink, "Hyperlink" },
                { (short) IVisio.VisSectionIndices.visSectionLayer, "Layer" },
                { (short) IVisio.VisSectionIndices.visSectionParagraph, "Paragraph" },
                { (short) IVisio.VisSectionIndices.visSectionProp, "Prop" },
                { (short) IVisio.VisSectionIndices.visSectionReviewer, "Reviewer" },
                { (short) IVisio.VisSectionIndices.visSectionScratch, "Scratch" },
                { (short) IVisio.VisSectionIndices.visSectionSmartTag, "SmartTag" },
                { (short) IVisio.VisSectionIndices.visSectionTab, "Tab" },
                { (short) IVisio.VisSectionIndices.visSectionTextField, "TextField" },
                { (short) IVisio.VisSectionIndices.visSectionUser, "User" },
                { (short) IVisio.VisSectionIndices.visSectionObject , "Object"}

            };

        private static short[] SectionObject_RowIndices = new short[]
                                                              {
(short) IVisio.VisRowIndices.visRowAlign,
(short) IVisio.VisRowIndices.visRowDoc,
(short) IVisio.VisRowIndices.visRowEvent,
(short) IVisio.VisRowIndices.visRowForeign,
(short) IVisio.VisRowIndices.visRowFill,
(short) IVisio.VisRowIndices.visRowMisc,
(short) IVisio.VisRowIndices.visRowGroup,

(short) IVisio.VisRowIndices.visRowImage,
(short) IVisio.VisRowIndices.visRowLine,
(short) IVisio.VisRowIndices.visRowMisc,
(short) IVisio.VisRowIndices.visRowXForm1D,
(short) IVisio.VisRowIndices.visRowPageLayout,
(short) IVisio.VisRowIndices.visRowPrintProperties,

(short) IVisio.VisRowIndices.visRowPage,
(short) IVisio.VisRowIndices.visRowParagraph,
(short) IVisio.VisRowIndices.visRowLock,

(short) IVisio.VisRowIndices.visRowRulerGrid,
(short) IVisio.VisRowIndices.visRowXFormOut,
(short) IVisio.VisRowIndices.visRowTextXForm,                                                                  

(short) IVisio.VisRowIndices.visRowText,
(short) IVisio.VisRowIndices.visRowStyle,
(short) IVisio.VisRowIndices.visRowShapeLayout,                                                                  

    };
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void CheckCellNames()
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

                var stops = new[]
                                {
                                    new VA.Text.TabStop(0.1, VA.Text.TabStopAlignment.Left),
                                    new VA.Text.TabStop(0.2, VA.Text.TabStopAlignment.Right)
                                };

                VA.Text.TextHelper.SetTabStops(shape1,stops);

            }

            System.Threading.Thread.Sleep(1000);

            foreach (short section_index in common_section_indices)
            {
                Debug.WriteLine(TryGetSectionName(section_index) ?? "UNKNOWN SECTION");
                Debug.WriteLine("--------------------");
                foreach (var cellinfo in EnumCellsInSection(shape1, section_index))
                {
                    Debug.WriteLine("{0} {1} : {2} {3} // (\"{4}\", {5})", cellinfo.RealName, cellinfo.SRC.ToString(), cellinfo.XName, cellinfo.XSRC.ToString(), cellinfo.Formula, cellinfo.Result);

                }
            }

            doc1.Close(true);
        }


        private string TryGetSectionName(short si)
        {
            if (section_to_name.ContainsKey((short)si))
            {
                return section_to_name[(short)si];
            }
            return null;
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
                    if (!xcellsrc.HasValue)
                    {
                        xcellsrc = new VA.ShapeSheet.SRC(-1, -1, -1);
                    }


                    var ci = new CellInfo();
                    ci.RealName = cell_name;
                    ci.SRC = cell_src;

                    ci.XName = "TBD";
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
