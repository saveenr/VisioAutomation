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
        public Rows Rows = new Rows();
        public Sections Sections = new Sections();

        public Dictionary<short, string> SectionToName;

        public ShapeSheetMetadata()
        {
            this.SectionToName = new Dictionary<short, string>();
            foreach (var section in this.Sections.Items)
            {
                this.SectionToName[section.EnumValue] = section.DisplayName;
            }

            this.CommonSections = new List<SectionDef>();
            this.CommonSections.Add(this.Sections.Action);
                this.CommonSections.Add( this.Sections.Annotation);
                this.CommonSections.Add( this.Sections.Character);
                this.CommonSections.Add( this.Sections.ConnectionPts);
                this.CommonSections.Add( this.Sections.Controls);
                this.CommonSections.Add( this.Sections.Hyperlink);
                this.CommonSections.Add( this.Sections.Layer);
                this.CommonSections.Add( this.Sections.Paragraph);
                this.CommonSections.Add( this.Sections.Prop);
                this.CommonSections.Add( this.Sections.Reviewer);
                this.CommonSections.Add( this.Sections.Scratch );
                this.CommonSections.Add( this.Sections.SmartTag);
                this.CommonSections.Add( this.Sections.Tab);
                this.CommonSections.Add( this.Sections.TextField);
                this.CommonSections.Add(this.Sections.User);
                this.CommonSections.Add( this.Sections.Object  );



        }


        public List<SectionDef> CommonSections;
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

    public class SectionDef
    {
        public readonly string DisplayName;
        public readonly string EnumName;
        public readonly short EnumValue;

        public SectionDef(string displayname, string enumname, IVisio.VisSectionIndices enumvalue)
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

    public class Rows
    {
            public RowDef Doc = new RowDef( "Doc", "visRowDoc" ,ROW.visRowDoc );
            public RowDef Event =new RowDef( "Event", "visRowEvent" ,ROW.visRowEvent );
            public RowDef Foreign =new RowDef( "Foreign", "visRowForeign" ,ROW.visRowForeign );
            public RowDef Fill = new RowDef( "Fill", "visRowFill" ,ROW.visRowFill );
            public RowDef Misc = new RowDef( "Misc", "visRowMisc" ,ROW.visRowMisc );
            public RowDef Group = new RowDef( "Group", "visRowGroup" ,ROW.visRowGroup );
            public RowDef Image = new RowDef( "Image", "visRowImage" ,ROW.visRowImage );
            public RowDef Line = new RowDef( "Line", "visRowLine" ,ROW.visRowLine );
            public RowDef XForm1D = new RowDef( "XForm1D", "visRowXForm1D" ,ROW.visRowXForm1D );
            public RowDef PageLayout = new RowDef( "PageLayout", "visRowPageLayout" ,ROW.visRowPageLayout );
            public RowDef PrintProperties = new RowDef( "PrintProperties", "visRowPrintProperties" ,ROW.visRowPrintProperties );
            public RowDef Page = new RowDef( "Page", "visRowPage" ,ROW.visRowPage );
            public RowDef Paragraph = new RowDef( "Paragraph", "visRowParagraph" ,ROW.visRowParagraph );
            public RowDef Lock = new RowDef( "Lock", "visRowLock" ,ROW.visRowLock );
            public RowDef RulerGrid = new RowDef( "RulerGrid", "visRowRulerGrid" ,ROW.visRowRulerGrid );
            public RowDef XFormOut = new RowDef( "XFormOut", "visRowXFormOut" ,ROW.visRowXFormOut );
            public RowDef TextXForm = new RowDef( "TextXForm", "visRowTextXForm" ,ROW.visRowTextXForm );
            public RowDef Text = new RowDef( "Text", "visRowText" ,ROW.visRowText );
            public RowDef Style = new RowDef( "Style", "visRowStyle" ,ROW.visRowStyle );
            public RowDef ShapeLayout = new RowDef("ShapeLayout", "visRowShapeLayout", ROW.visRowShapeLayout);

            public IEnumerable<RowDef> Items
            {
                get
                {
                    yield return this.Doc;
                    yield return this.Event;
                    yield return this.Foreign;
                    yield return this.Fill;
                    yield return this.Misc;
                    yield return this.Group;
                    yield return this.Image;
                    yield return this.Line;
                    yield return this.XForm1D;
                    yield return this.PageLayout;
                    yield return this.PrintProperties;
                    yield return this.Page;
                    yield return this.Paragraph;
                    yield return this.Lock;
                    yield return this.RulerGrid;
                    yield return this.XFormOut;
                    yield return this.TextXForm;
                    yield return this.Text;
                    yield return this.Style;
                    yield return this.ShapeLayout;                
                }
            }
            
    }

    public class Sections
    {
        public SectionDef Action = new SectionDef( "Action", "visSectionAction" ,SEC.visSectionAction );
        public SectionDef Annotation = new SectionDef( "Annotation", "visSectionAnnotation" ,SEC.visSectionAnnotation );
        public SectionDef Character = new SectionDef( "Character", "visSectionCharacter" ,SEC.visSectionCharacter );
        public SectionDef ConnectionPts = new SectionDef( "ConnectionPts", "visSectionConnectionPts" ,SEC.visSectionConnectionPts );
        public SectionDef Controls = new SectionDef( "Controls", "visSectionControls" ,SEC.visSectionControls );
        public SectionDef Hyperlink = new SectionDef( "Hyperlink", "visSectionHyperlink" ,SEC.visSectionHyperlink );
        public SectionDef Layer = new SectionDef( "Layer", "visSectionLayer" ,SEC.visSectionLayer );
        public SectionDef Paragraph = new SectionDef( "Paragraph", "visSectionParagraph" ,SEC.visSectionParagraph );
        public SectionDef Prop = new SectionDef( "Prop", "visSectionProp" ,SEC.visSectionProp );
        public SectionDef Reviewer = new SectionDef( "Reviewer", "visSectionReviewer" ,SEC.visSectionReviewer );
        public SectionDef Scratch = new SectionDef( "Scratch", "visSectionScratch" ,SEC.visSectionScratch );
        public SectionDef SmartTag = new SectionDef( "SmartTag", "visSectionSmartTag" ,SEC.visSectionSmartTag );
        public SectionDef Tab = new SectionDef( "Tab", "visSectionTab" ,SEC.visSectionTab );
        public SectionDef TextField = new SectionDef( "TextField", "visSectionTextField" ,SEC.visSectionTextField );
        public SectionDef User = new SectionDef( "User", "visSectionUser" ,SEC.visSectionUser );
        public SectionDef Object = new SectionDef( "Object", "visSectionObject" ,SEC.visSectionObject );

        public IEnumerable<SectionDef> Items
        {
            get
            {
                yield return this.Action;
                yield return this.Annotation;
                yield return this.Character;
                yield return this.ConnectionPts;
                yield return this.Controls;
                yield return this.Hyperlink;
                yield return this.Layer;
                yield return this.Paragraph;
                yield return this.Prop;
                yield return this.Reviewer;
                yield return this.Scratch;
                yield return this.SmartTag;
                yield return this.Tab;
                yield return this.TextField;
                yield return this.User;
                yield return this.Object;
            }
        }
    }
}
