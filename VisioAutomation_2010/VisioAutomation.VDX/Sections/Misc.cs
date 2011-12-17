using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;

namespace VisioAutomation.VDX.Sections
{
    public class Misc
    {
        public IntCell NoObjHandles = new IntCell();
        public IntCell NonPrinting = new IntCell();
        public IntCell NoCtlHandles = new IntCell();
        public IntCell NoAlignBox = new IntCell();

        public IntCell UpdateAlignBox = new IntCell();
        public IntCell HideText = new IntCell();
        public IntCell DynFeedback = new IntCell();
        public IntCell GlueType = new IntCell();
        public IntCell WalkPreference = new IntCell();

        public DoubleCell BegTrigger = new DoubleCell();
        public DoubleCell EndTrigger = new DoubleCell();

        public IntCell ObjType = new IntCell();
        public IntCell Comment = new IntCell();
        public IntCell IsDropSource = new IntCell();
        public IntCell NoLiveDynamics = new IntCell();
        public IntCell LocalizeMerge = new IntCell();

        public IntCell Calendar = new IntCell();
        public IntCell LangID = new IntCell();
        public DoubleCell ShapeKeywords = new DoubleCell();
        public IntCell DropOnPageScale = new IntCell();

        public void AddToElement(System.Xml.Linq.XElement parent)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("Misc");
            el.Add(this.NoObjHandles.ToXml("NoObjHandles"));
            el.Add(this.NonPrinting.ToXml("NonPrinting"));
            el.Add(this.NoCtlHandles.ToXml("NoCtlHandles"));
            el.Add(this.NoAlignBox.ToXml("NoAlignBox"));

            el.Add(this.UpdateAlignBox.ToXml("UpdateAlignBox"));
            el.Add(this.HideText.ToXml("HideText"));
            el.Add(this.DynFeedback.ToXml("DynFeedback"));
            el.Add(this.GlueType.ToXml("GlueType"));
            el.Add(this.WalkPreference.ToXml("WalkPreference"));

            el.Add(this.BegTrigger.ToXml("BegTrigger"));
            el.Add(this.EndTrigger.ToXml("EndTrigger"));

            el.Add(this.ObjType.ToXml("ObjType"));
            el.Add(this.Comment.ToXml("Comment"));
            el.Add(this.IsDropSource.ToXml("IsDropSource"));
            el.Add(this.NoLiveDynamics.ToXml("NoLiveDynamics"));
            el.Add(this.LocalizeMerge.ToXml("LocalizeMerge"));

            el.Add(this.Calendar.ToXml("Calendar"));
            el.Add(this.LangID.ToXml("LangID"));
            el.Add(this.ShapeKeywords.ToXml("ShapeKeywords"));
            el.Add(this.DropOnPageScale.ToXml("DropOnPageScale"));

            parent.Add(el);
        }
    }
}