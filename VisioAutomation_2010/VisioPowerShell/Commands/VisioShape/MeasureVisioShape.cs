using System.Linq;
using SMA = System.Management.Automation;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS=VisioAutomation.ShapeSheet;

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

            int n = 0;
            foreach (var row in cellqueryresult)
            {
                var dim = new Models.ShapeDimensions();

                dim.ShapeID = shapeids[n];

                dim.XFormAngle = row[col_XFormAngle];
                dim.XFormWidth = row[col_XFormWidth];
                dim.XFormHeight = row[col_XFormHeight];
                dim.XFormLocPinX = row[col_XFormLocPinX];
                dim.XFormLocPinY = row[col_XFormLocPinY];
                dim.XFormPinX = row[col_XFormPinX];
                dim.XFormPinY = row[col_XFormPinY];

                dim.OneDBeginX = row[col_OneDBeginX];
                dim.OneDBeginY = row[col_OneDBeginY];
                dim.OneDEndX = row[col_OneDEndX];
                dim.OneDEndY = row[col_OneDEndY];

                this.WriteObject(dim);

                n++;
            }
        }
    }
}