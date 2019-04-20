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
        public IVisio.Page [] Page;

        /*
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TreatUndirectedAsBidirectional { get; set; }
        */

        protected override void ProcessRecord()
        {

            var targetpages = new VisioScripting.TargetPages(this.Page).Resolve(this.Client);

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


namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Measure, Nouns.VisioShape)]
    public class MeasureVisioShape: VisioCmdlet
    {

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {

            var targetshapes = new VisioScripting.TargetShapes(this.Shape).Resolve(this.Client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }
           

            var query = new VASS.Query.CellQuery();

            var col_XFormAngle = query.Columns.Add(VASS.SrcConstants.XFormAngle, nameof(VASS.SrcConstants.XFormAngle));
            var col_XFormHeight = query.Columns.Add(VASS.SrcConstants.XFormHeight, nameof(VASS.SrcConstants.XFormHeight));
            var col_XFormWidth = query.Columns.Add(VASS.SrcConstants.XFormWidth, nameof(VASS.SrcConstants.XFormWidth));
            var col_XFormLocPinX = query.Columns.Add(VASS.SrcConstants.XFormLocPinX, nameof(VASS.SrcConstants.XFormLocPinX));
            var col_XFormLocPinY = query.Columns.Add(VASS.SrcConstants.XFormLocPinY, nameof(VASS.SrcConstants.XFormLocPinY));
            var col_XFormPinX = query.Columns.Add(VASS.SrcConstants.XFormPinX, nameof(VASS.SrcConstants.XFormPinX));
            var col_XFormPinY = query.Columns.Add(VASS.SrcConstants.XFormPinY, nameof(VASS.SrcConstants.XFormPinY));

            var col_OneDBeginX = query.Columns.Add(VASS.SrcConstants.OneDBeginX, nameof(VASS.SrcConstants.OneDBeginX));
            var col_OneDBeginY = query.Columns.Add(VASS.SrcConstants.OneDBeginY, nameof(VASS.SrcConstants.OneDBeginY));
            var col_OneDEndX = query.Columns.Add(VASS.SrcConstants.OneDEndX, nameof(VASS.SrcConstants.OneDEndX));
            var col_OneDEndY = query.Columns.Add(VASS.SrcConstants.OneDEndY, nameof(VASS.SrcConstants.OneDEndY));

            var page = targetshapes.Shapes[0].ContainingPage;
            var shapeids = VisioAutomation.ShapeIDPairs.FromShapes(targetshapes.Shapes).Select(i => i.ShapeID).ToList();
            var cellqueryresult = query.GetResults<double>(page,shapeids);
            foreach (var row in cellqueryresult)
            {
                var dim = new Models.ShapeDimensions();
                dim.XFormAngle = row[col_XFormAngle];
                dim.XFormWidth = row[col_XFormWidth];
                dim.XFormHeight = row[col_XFormHeight];
                dim.XFormLocPinX = row[col_XFormLocPinX];
                dim.XFormLocPinY = row[col_XFormLocPinY];
                dim.XFormPinX = row[col_XFormPinX];
                dim.XFormPinY = row[col_XFormPinY];
                this.WriteObject(dim);
            }
        }
    }
}