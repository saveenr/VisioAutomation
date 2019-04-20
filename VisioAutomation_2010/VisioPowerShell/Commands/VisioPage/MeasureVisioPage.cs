using System.Linq;
using SMA = System.Management.Automation;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Measure, Nouns.VisioPage)]
    public class MeasureVisioPage : VisioCmdlet
    {

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page [] Pages;

        /*
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TreatUndirectedAsBidirectional { get; set; }
        */

        protected override void ProcessRecord()
        {

            var targetpages = new VisioScripting.TargetPages(this.Pages).Resolve(this.Client);

            if (targetpages.Pages.Count < 1)
            {
                return;
            }

            var query = new VASS.Query.CellQuery();
            var col_PageHeight = query.Columns.Add(VASS.SrcConstants.PageHeight, nameof(VASS.SrcConstants.PageHeight));
            var col_PageWidth = query.Columns.Add(VASS.SrcConstants.PageWidth, nameof(VASS.SrcConstants.PageWidth));
            var col_PrintBottomMargin = query.Columns.Add(VASS.SrcConstants.PrintBottomMargin, nameof(VASS.SrcConstants.PrintBottomMargin));
            var col_PrintTopMargin = query.Columns.Add(VASS.SrcConstants.PrintTopMargin, nameof(VASS.SrcConstants.PrintTopMargin));
            var col_PrintLeftMargin = query.Columns.Add(VASS.SrcConstants.PrintLeftMargin, nameof(VASS.SrcConstants.PrintLeftMargin));
            var col_PrintRightMargin = query.Columns.Add(VASS.SrcConstants.PrintRightMargin, nameof(VASS.SrcConstants.PrintRightMargin));

            foreach (var page in targetpages.Pages)
            {
                var pd = new Models.PageDimensions();

                var cellqueryresult = query.GetResults<double>(page.PageSheet);
                var row = cellqueryresult[0];
                pd.PageHeight = row[col_PageHeight];
                pd.PageWidth = row[col_PageWidth];
                pd.PrintBottomMargin = row[col_PrintBottomMargin];
                pd.PrintLeftMargin = row[col_PrintLeftMargin];
                pd.PrintRightMargin = row[col_PrintRightMargin];
                pd.PrintTopMargin = row[col_PrintTopMargin];

                this.WriteObject(pd);

            }
        }

        private void foo()
        {
            /*

            var targetpage = new VisioScripting.TargetPage(this.Page);

            var options = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options.NoArrowsHandling = VA.DocumentAnalysis.NoArrowsHandling.ExcludeEdge;

            options.DirectionSource = VA.DocumentAnalysis.DirectionSource.UseConnectorArrows;

            options.NoArrowsHandling = this.TreatUndirectedAsBidirectional ?
                VA.DocumentAnalysis.NoArrowsHandling.TreatEdgeAsBidirectional
                : VA.DocumentAnalysis.NoArrowsHandling.ExcludeEdge;
            var edges = this.Client.Connection.GetDirectedEdgesOnPage(targetpage, options);
            this.WriteObject(edges, false);
                */
        }
    }
}

