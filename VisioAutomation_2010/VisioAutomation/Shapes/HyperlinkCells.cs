using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class HyperlinkCells : CellGroup
    {
        public CellValueLiteral Address { get; set; }
        public CellValueLiteral Description { get; set; }
        public CellValueLiteral ExtraInfo { get; set; }
        public CellValueLiteral Frame { get; set; }
        public CellValueLiteral SortKey { get; set; }
        public CellValueLiteral SubAddress { get; set; }
        public CellValueLiteral NewWindow { get; set; }
        public CellValueLiteral Default { get; set; }
        public CellValueLiteral Invisible { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.HyperlinkAddress, this.Address);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkDescription, this.Description);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkExtraInfo, this.ExtraInfo);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkFrame, this.Frame);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkSortKey, this.SortKey);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkSubAddress, this.SubAddress);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkNewWindow, this.NewWindow);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkDefault, this.Default);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkInvisible, this.Invisible);
            }
        }
    }
}