using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Connections
{
    class ConnectionPointQuery : VA.ShapeSheet.Query.SectionQuery
    {
        public VA.ShapeSheet.Query.SectionQueryColumn DirX { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn DirY { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Type { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn X { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Y { get; set; }

        public ConnectionPointQuery() :
            base(IVisio.VisSectionIndices.visSectionConnectionPts)
        {
            DirX = this.AddColumn(IVisio.VisCellIndices.visCnnctDirX,"DirX");
            DirY = this.AddColumn(IVisio.VisCellIndices.visCnnctDirY,"DirY");
            Type = this.AddColumn(IVisio.VisCellIndices.visCnnctType,"Type");
            X = this.AddColumn(IVisio.VisCellIndices.visX, "X");
            Y = this.AddColumn(IVisio.VisCellIndices.visY, "Y");
        }
    }
}