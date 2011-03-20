using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.CustomProperties
{
    class CustomPropertyQuery : VA.ShapeSheet.Query.SectionQuery
    {
        public VA.ShapeSheet.Query.SectionQueryColumn SortKey { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Ask { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Calendar { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Format { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Invis { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Label { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn LangID { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Prompt { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Value { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Type { get; set; }

        public CustomPropertyQuery() :
            base(IVisio.VisSectionIndices.visSectionProp)
        {
            SortKey = this.AddColumn(IVisio.VisCellIndices.visCustPropsSortKey, "SortKey");
            Ask = this.AddColumn(IVisio.VisCellIndices.visCustPropsAsk, "Ask");
            Calendar = this.AddColumn(IVisio.VisCellIndices.visCustPropsCalendar, "Calendar");
            Format = this.AddColumn(IVisio.VisCellIndices.visCustPropsFormat, "Format");
            Invis = this.AddColumn(IVisio.VisCellIndices.visCustPropsInvis, "Invis");
            Label = this.AddColumn(IVisio.VisCellIndices.visCustPropsLabel, "Label");
            LangID = this.AddColumn(IVisio.VisCellIndices.visCustPropsLangID, "LangID");
            Prompt = this.AddColumn(IVisio.VisCellIndices.visCustPropsPrompt, "Prompt");
            Type = this.AddColumn(IVisio.VisCellIndices.visCustPropsType, "Type");
            Value = this.AddColumn(IVisio.VisCellIndices.visCustPropsValue, "Value");
        }
    }
}