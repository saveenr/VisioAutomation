using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections
{
    public class Protection
    {
        public BoolCell Width = new BoolCell();
        public BoolCell Height = new BoolCell();
        public BoolCell MoveX = new BoolCell();
        public BoolCell MoveY = new BoolCell();
        public BoolCell Aspect = new BoolCell();
        public BoolCell Delete = new BoolCell();
        public BoolCell Begin = new BoolCell();
        public BoolCell Rotate = new BoolCell();
        public BoolCell Crop = new BoolCell();
        public BoolCell VtxEdit = new BoolCell();

        public BoolCell TextEdit = new BoolCell();
        public BoolCell Format = new BoolCell();
        public BoolCell Group = new BoolCell();
        public BoolCell CalcWH = new BoolCell();
        public BoolCell Select = new BoolCell();
        public BoolCell CustProp = new BoolCell();

        //<vx:Protection xmlns:vx="http://schemas.microsoft.com/visio/2006/extension">
        //<vx:LockFromGroupFormat>0</vx:LockFromGroupFormat>
        //<vx:LockThemeColors>0</vx:LockThemeColors>
        //<vx:LockThemeEffects>0</vx:LockThemeEffects>
        //</vx:Protection>

        public BoolCell FromGroupFormat = new BoolCell();
        public BoolCell ThemeColors = new BoolCell();
        public BoolCell ThemeEffects = new BoolCell();

        public void AddToElement(SXL.XElement parent)
        {
            var el1 = XMLUtil.CreateVisioSchema2003Element("Protection");
            el1.Add(this.Width.ToXml("LockWidth"));
            el1.Add(this.Height.ToXml("LockHeight"));

            el1.Add(this.MoveX.ToXml("LockMoveX"));
            el1.Add(this.MoveY.ToXml("LockMoveY"));

            el1.Add(this.Aspect.ToXml("LockAspect"));
            el1.Add(this.Delete.ToXml("LockDelete"));

            el1.Add(this.Begin.ToXml("LockBegin"));
            el1.Add(this.Rotate.ToXml("LockRotate"));

            el1.Add(this.Crop.ToXml("LockCrop"));
            el1.Add(this.VtxEdit.ToXml("LockVtxEdit"));

            el1.Add(this.TextEdit.ToXml("LockTextEdit"));
            el1.Add(this.Format.ToXml("LockFormat"));

            el1.Add(this.Group.ToXml("LockGroup"));
            el1.Add(this.CalcWH.ToXml("LockCalcWH"));

            el1.Add(this.Select.ToXml("LockSelect"));
            el1.Add(this.CustProp.ToXml("LockCustProp"));

            parent.Add(el1);

            var el2 = XMLUtil.CreateVisioSchema2006Element("Protection");
            el2.Add(this.FromGroupFormat.ToXml2006("LockFromGroupFormat"));
            el2.Add(this.ThemeColors.ToXml2006("LockThemeColors"));
            el2.Add(this.ThemeEffects.ToXml2006("LockThemeEffects"));
            parent.Add(el2);
        }

        public void SetAll(bool v)
        {
            Width.Result = v;
            Height.Result = v;
            MoveX.Result = v;
            MoveY.Result = v;
            Aspect.Result = v;
            Delete.Result = v;
            Begin.Result = v;
            Rotate.Result = v;
            Crop.Result = v;
            VtxEdit.Result = v;
            TextEdit.Result = v;
            Format.Result = v;
            Group.Result = v;
            CalcWH.Result = v;
            Select.Result = v;
            CustProp.Result = v;
        }
    }
}