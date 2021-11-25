using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
namespace VisioScripting.Models
{
    public class PageDimensions
    {
        public int PageID;

        public double PageHeight;
        public double PageWidth;

        public double PrintLeftMargin;
        public double PrintRightMargin;

        public double PrintTopMargin;
        public double PrintBottomMargin;


        public static List<PageDimensions> Get_PageDimensions(IList<IVisio.Page> pages)
        {
            var list_pagedim = new List<VisioScripting.Models.PageDimensions>(pages.Count);

            var query = new VASS.Query.CellQuery();
            var col_PageHeight = query.Columns.Add(VisioAutomation.Core.SrcConstants.PageHeight, nameof(VisioAutomation.Core.SrcConstants.PageHeight));
            var col_PageWidth = query.Columns.Add(VisioAutomation.Core.SrcConstants.PageWidth, nameof(VisioAutomation.Core.SrcConstants.PageWidth));
            var col_PrintBottomMargin =
                query.Columns.Add(VisioAutomation.Core.SrcConstants.PrintBottomMargin, nameof(VisioAutomation.Core.SrcConstants.PrintBottomMargin));
            var col_PrintTopMargin =
                query.Columns.Add(VisioAutomation.Core.SrcConstants.PrintTopMargin, nameof(VisioAutomation.Core.SrcConstants.PrintTopMargin));
            var col_PrintLeftMargin =
                query.Columns.Add(VisioAutomation.Core.SrcConstants.PrintLeftMargin, nameof(VisioAutomation.Core.SrcConstants.PrintLeftMargin));
            var col_PrintRightMargin =
                query.Columns.Add(VisioAutomation.Core.SrcConstants.PrintRightMargin, nameof(VisioAutomation.Core.SrcConstants.PrintRightMargin));


            foreach (var page in pages)
            {
                var pagedim = new VisioScripting.Models.PageDimensions();

                pagedim.PageID = page.ID;

                var cellqueryresult = query.GetResults<double>(page.PageSheet);
                var row = cellqueryresult[0];
                pagedim.PageHeight = row[col_PageHeight];
                pagedim.PageWidth = row[col_PageWidth];
                pagedim.PrintBottomMargin = row[col_PrintBottomMargin];
                pagedim.PrintLeftMargin = row[col_PrintLeftMargin];
                pagedim.PrintRightMargin = row[col_PrintRightMargin];
                pagedim.PrintTopMargin = row[col_PrintTopMargin];

                list_pagedim.Add(pagedim);
            }

            return list_pagedim;
        }

    }
}