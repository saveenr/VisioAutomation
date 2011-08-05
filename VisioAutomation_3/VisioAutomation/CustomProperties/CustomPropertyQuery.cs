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
            SortKey = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_SortKey, "SortKey");
            Ask = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Ask, "Ask");
            Calendar = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Calendar, "Calendar");
            Format = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Format, "Format");
            Invis = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Invisible, "Invis");
            Label = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Label, "Label");
            LangID = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_LangID, "LangID");
            Prompt = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Prompt, "Prompt");
            Type = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Value, "Type");
            Value = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Value, "Value");
        }
    }
}