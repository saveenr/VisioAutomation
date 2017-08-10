using VisioAutomation.ShapeSheet;

namespace VisioScripting.Models
{
    public class CellSrcDictionary : NamedSrcDictionary
    {
        private static CellSrcDictionary shape_cellmap;
        private static CellSrcDictionary page_cellmap;

        public static CellSrcDictionary GetCellMapForShapes()
        {
            if (CellSrcDictionary.shape_cellmap == null)
            {
                CellSrcDictionary.shape_cellmap = new CellSrcDictionary();

                var pagecells = new VisioScripting.Models.PageCells();
                foreach (var t in pagecells.GetSrcFormulaPairs())
                {
                    CellSrcDictionary.shape_cellmap[t.Name] = t.Src;
                }
            }
            return CellSrcDictionary.shape_cellmap;
        }

        public static CellSrcDictionary GetCellMapForPages()
        {
            if (CellSrcDictionary.page_cellmap == null)
            {
                CellSrcDictionary.page_cellmap = new CellSrcDictionary();

                var shapecells = new VisioScripting.Models.ShapeCells();
                foreach (var t in shapecells.GetSrcFormulaPairs())
                {
                    CellSrcDictionary.shape_cellmap[t.Name] = t.Src;
                }
            }
            return CellSrcDictionary.page_cellmap;
        }
    }
}


    
