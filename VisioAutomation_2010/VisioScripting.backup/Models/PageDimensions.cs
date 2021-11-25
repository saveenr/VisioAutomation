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
            var col_PageHeight = query.Columns.Add(VASS.SrcConstants.PageHeight, nameof(VASS.SrcConstants.PageHeight));
            var col_PageWidth = query.Columns.Add(VASS.SrcConstants.PageWidth, nameof(VASS.SrcConstants.PageWidth));
            var col_PrintBottomMargin =
                query.Columns.Add(VASS.SrcConstants.PrintBottomMargin, nameof(VASS.SrcConstants.PrintBottomMargin));
            var col_PrintTopMargin =
                query.Columns.Add(VASS.SrcConstants.PrintTopMargin, nameof(VASS.SrcConstants.PrintTopMargin));
            var col_PrintLeftMargin =
                query.Columns.Add(VASS.SrcConstants.PrintLeftMargin, nameof(VASS.SrcConstants.PrintLeftMargin));
            var col_PrintRightMargin =
                query.Columns.Add(VASS.SrcConstants.PrintRightMargin, nameof(VASS.SrcConstants.PrintRightMargin));


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