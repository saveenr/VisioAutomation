using Microsoft.Office.Interop.Visio;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioInterop
{
    public static class ListExtensions
    {
        public static void Set3(this short[] list, int index, short section, short row, short cell)
        {
            list[(index * 3) + 0] = section;
            list[(index * 3) + 1] = row;
            list[(index * 3) + 2] = cell;
        }

        public static void Set4(this short[] list, int index, short shapeid, short section, short row, short cell)
        {
            list[(index * 4) + 0] = shapeid;
            list[(index * 4) + 1] = section;
            list[(index * 4) + 2] = row;
            list[(index * 4) + 3] = cell;
        }

    }

    public static class Util
    {
        public static Shape CreateStandardShape(Page page)
        {
            var shape = page.DrawRectangle(1, 1, 4, 3);
            var cell_width = shape.CellsU["Width"];
            var cell_height = shape.CellsU["Height"];
            cell_width.Formula = "=(1.0+2.5)";
            cell_height.Formula = "=(0.0+1.5)";
            return shape;
        }

        public static Page CreateStandardPage(Document doc, string pagename)
        {
            var pages = doc.Pages;
            var page = pages.Add();
            page.NameU = pagename;
            return page;
        }

        public static CellSRC Cell_Width = new CellSRC((short)VisSectionIndices.visSectionObject,
                                               (short)VisRowIndices.visRowXFormOut,
                                               (short)VisCellIndices.visXFormWidth);

        public static CellSRC Cell_Height = new CellSRC((short)VisSectionIndices.visSectionObject,
                                                        (short)VisRowIndices.visRowXFormOut,
                                                        (short)VisCellIndices.visXFormHeight);

        public static SIDSRCGetFormulaItem[] Create_PageGetFormulas_Request(Shape shape)
        {
            var shapeid = (short) shape.ID;
            return new[]
                       {
                           new SIDSRCGetFormulaItem(shapeid, Cell_Width),
                           new SIDSRCGetFormulaItem(shapeid, Cell_Height)
                       };
        }

        public static SIDSRCSetResultItem[] Create_PageSetResults_Request(Shape shape)
        {
            var shapeid = (short) shape.ID;
            return new[]
                       {
                           new SIDSRCSetResultItem(shapeid,Cell_Width,  8.0, (short) VisUnitCodes.visNoCast ),
                           new SIDSRCSetResultItem(shapeid,Cell_Height, 1.0, (short) VisUnitCodes.visNoCast)
                       };
        }

        public static SIDSRCSetFormulaItem[] Create_PageSetFormulas_Request(Shape shape)
        {
            var shapeid = (short) shape.ID;
            return new[]
                       {
                           new SIDSRCSetFormulaItem(shapeid,Cell_Width,  "1.3"),
                           new SIDSRCSetFormulaItem(shapeid,Cell_Height, "1.7")
                       };
        }

        public static SIDSRCGetResultItem[] Create_PageGetResults_Request(Shape shape)
        {
            var shapeid = (short) shape.ID;

            return new[]
                       {
                           new SIDSRCGetResultItem(shapeid,Cell_Width,  (short) VisUnitCodes.visNoCast),
                           new SIDSRCGetResultItem(shapeid,Cell_Height, (short) VisUnitCodes.visNoCast)
                       };
        }

        public static SRCGetResultItem[] Create_ShapeGetResults_Request()
        {
            return new[]
                       {
                           new SRCGetResultItem(Cell_Width,  (short) VisUnitCodes.visNoCast),
                           new SRCGetResultItem(Cell_Height, (short) VisUnitCodes.visNoCast)
                       };
        }

        public static SRCSetFormulaItem[] Create_ShapeSetFormulas_Request()
        {
            return new[]
                       {
                           new SRCSetFormulaItem(Cell_Width,  "1.3"),
                           new SRCSetFormulaItem(Cell_Height, "7.1")
                       };
        }

        public static SRCSetResultItem[] Create_ShapeSetResults_Request()
        {
            return new[]
                       {
                           new SRCSetResultItem(Cell_Width,  8.0, (short) VisUnitCodes.visNoCast ),
                           new SRCSetResultItem(Cell_Height, 1.0, (short) VisUnitCodes.visNoCast)
                       };
        }

        public static SRCGetFormulaItem[] Create_ShapeGetFormulas_Request()
        {
            return new[]
                       {
                           new VisioInterop.SRCGetFormulaItem(Cell_Width ),
                           new VisioInterop.SRCGetFormulaItem(Cell_Height)
                       };
        }
    }
}