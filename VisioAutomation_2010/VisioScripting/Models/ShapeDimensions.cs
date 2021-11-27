using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    public class ShapeDimensions
    {
        public int ShapeID;

        public double XFormAngle;
        public double XFormHeight;
        public double XFormLocPinX;
        public double XFormLocPinY;
        public double XFormPinX;
        public double XFormPinY;
        public double XFormWidth;

        public double OneDBeginX;
        public double OneDBeginY;
        public double OneDEndX;
        public double OneDEndY;


        public static List<ShapeDimensions> Get_ShapeDimensions(IVisio.Page page, List<int> shapeids)
        {
            var query = new VASS.Query.CellQuery();

            var col_XFormAngle = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormAngle);
            var col_XFormHeight = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormHeight);
            var col_XFormWidth = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormWidth);
            var col_XFormLocPinX = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormLocPinX);
            var col_XFormLocPinY = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormLocPinY);
            var col_XFormPinX = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormPinX);
            var col_XFormPinY = query.Columns.Add(VisioAutomation.Core.SrcConstants.XFormPinY);

            var col_OneDBeginX = query.Columns.Add(VisioAutomation.Core.SrcConstants.OneDBeginX);
            var col_OneDBeginY = query.Columns.Add(VisioAutomation.Core.SrcConstants.OneDBeginY);
            var col_OneDEndX = query.Columns.Add(VisioAutomation.Core.SrcConstants.OneDEndX);
            var col_OneDEndY = query.Columns.Add(VisioAutomation.Core.SrcConstants.OneDEndY);

            var cellqueryresult = query.GetResults<double>(page, shapeids);

            var list_shapedim = new List<VisioScripting.Models.ShapeDimensions>(shapeids.Count);
            int n = 0;
            foreach (var row in cellqueryresult)
            {
                var shapedim = new VisioScripting.Models.ShapeDimensions();

                shapedim.ShapeID = shapeids[n];

                shapedim.XFormAngle = row[col_XFormAngle];
                shapedim.XFormWidth = row[col_XFormWidth];
                shapedim.XFormHeight = row[col_XFormHeight];
                shapedim.XFormLocPinX = row[col_XFormLocPinX];
                shapedim.XFormLocPinY = row[col_XFormLocPinY];
                shapedim.XFormPinX = row[col_XFormPinX];
                shapedim.XFormPinY = row[col_XFormPinY];

                shapedim.OneDBeginX = row[col_OneDBeginX];
                shapedim.OneDBeginY = row[col_OneDBeginY];
                shapedim.OneDEndX = row[col_OneDEndX];
                shapedim.OneDEndY = row[col_OneDEndY];

                list_shapedim.Add(shapedim);

                n++;
            }

            return list_shapedim;
        }
    }
}