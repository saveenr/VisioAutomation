using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Diagnostics;
using SEC = Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW = Microsoft.Office.Interop.Visio.VisRowIndices;
namespace TestVisioAutomation
{
    public class ShapeSheetMetadata
    {
        public ShapeSheetMetadata()
        {
            
        }
        public short[] CommonSectionIndices = new[] 
            {     
                (short) SEC.visSectionAction, 
                (short) SEC.visSectionAnnotation,
                (short) SEC.visSectionCharacter,
                (short) SEC.visSectionConnectionPts,
                (short) SEC.visSectionControls,
                (short) SEC.visSectionExport, 
                (short) SEC.visSectionHyperlink,
                (short) SEC.visSectionLayer, 
                (short) SEC.visSectionParagraph,
                (short) SEC.visSectionProp, 
                (short) SEC.visSectionReviewer,
                (short) SEC.visSectionScratch, 
                (short) SEC.visSectionSmartTag,
                (short) SEC.visSectionTab, 
                (short) SEC.visSectionTextField,
                (short) SEC.visSectionUser, 
                (short) SEC.visSectionObject  
            };

        public Dictionary<short, string> SectionToName = new Dictionary<short, string>
            {
                { (short) SEC.visSectionAction, "Action" },
                { (short) SEC.visSectionAnnotation, "Annotation" },
                { (short) SEC.visSectionCharacter, "Character" },
                { (short) SEC.visSectionConnectionPts, "ConnectionPts" },
                { (short) SEC.visSectionControls, "Controls" },
                { (short) SEC.visSectionHyperlink, "Hyperlink" },
                { (short) SEC.visSectionLayer, "Layer" },
                { (short) SEC.visSectionParagraph, "Paragraph" },
                { (short) SEC.visSectionProp, "Prop" },
                { (short) SEC.visSectionReviewer, "Reviewer" },
                { (short) SEC.visSectionScratch, "Scratch" },
                { (short) SEC.visSectionSmartTag, "SmartTag" },
                { (short) SEC.visSectionTab, "Tab" },
                { (short) SEC.visSectionTextField, "TextField" },
                { (short) SEC.visSectionUser, "User" },
                { (short) SEC.visSectionObject , "Object"}

            };

        public RowDef[] SectionObject_Rows = new RowDef[]
        {
            new RowDef( "Align", "visRowAlign" ,ROW.visRowAlign ),
            new RowDef( "Doc", "visRowDoc" ,ROW.visRowDoc ),
            new RowDef( "Event", "visRowEvent" ,ROW.visRowEvent ),
            new RowDef( "Foreign", "visRowForeign" ,ROW.visRowForeign ),
            new RowDef( "Fill", "visRowFill" ,ROW.visRowFill ),
            new RowDef( "Misc", "visRowMisc" ,ROW.visRowMisc ),
            new RowDef( "Group", "visRowGroup" ,ROW.visRowGroup ),
            new RowDef( "Image", "visRowImage" ,ROW.visRowImage ),
            new RowDef( "Line", "visRowLine" ,ROW.visRowLine ),
            new RowDef( "Misc", "visRowMisc" ,ROW.visRowMisc ),
            new RowDef( "XForm1D", "visRowXForm1D" ,ROW.visRowXForm1D ),
            new RowDef( "PageLayout", "visRowPageLayout" ,ROW.visRowPageLayout ),
            new RowDef( "PrintProperties", "visRowPrintProperties" ,ROW.visRowPrintProperties ),
            new RowDef( "Page", "visRowPage" ,ROW.visRowPage ),
            new RowDef( "Paragraph", "visRowParagraph" ,ROW.visRowParagraph ),
            new RowDef( "Lock", "visRowLock" ,ROW.visRowLock ),
            new RowDef( "RulerGrid", "visRowRulerGrid" ,ROW.visRowRulerGrid ),
            new RowDef( "XFormOut", "visRowXFormOut" ,ROW.visRowXFormOut ),
            new RowDef( "TextXForm", "visRowTextXForm" ,ROW.visRowTextXForm ),
            new RowDef( "Text", "visRowText" ,ROW.visRowText ),
            new RowDef( "Style", "visRowStyle" ,ROW.visRowStyle ),
            new RowDef( "ShapeLayout", "visRowShapeLayout" ,ROW.visRowShapeLayout )
        };
    }

    public class RowDef
    {
        public readonly string DisplayName;
        public readonly string EnumName;
        public readonly short EnumValue;

        public RowDef(string displayname, string enumname, IVisio.VisRowIndices enumvalue)
        {
            this.DisplayName = displayname;
            this.EnumName = enumname;
            this.EnumValue = (short) enumvalue;
        }
    }
    public class CellInfo
    {
        public string RealName;
        public VisioAutomation.ShapeSheet.SRC SRC;
        public string XName;
        public VisioAutomation.ShapeSheet.SRC XSRC;
        public string Formula;
        public double Result;

    }
}
