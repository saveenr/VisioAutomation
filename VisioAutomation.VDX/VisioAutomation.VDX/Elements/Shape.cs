using System.Xml.Linq;
using System.Collections.Generic;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.Internal.Extensions;
using VA=VisioAutomation;
using System.Linq;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Elements
{
    public class Shape : Node
    {
        public VA.VDX.Sections.XForm XForm = new VA.VDX.Sections.XForm();
        public VA.VDX.Sections.XForm1D XForm1D;
        public VA.VDX.Sections.Protection Protection;
        public VA.VDX.Sections.Fill Fill;
        public Line Line;
        public VA.VDX.Sections.TextXForm TextXForm;
        public VA.VDX.Sections.Misc Misc;
        public VA.VDX.Sections.Event Event;
        public VA.VDX.Sections.TextBlock TextBlock;
        public VA.VDX.Sections.Layout Layout;
        public List<VA.VDX.Sections.Char> CharFormats;
        public List<Sections.ParagraphFormat> ParaFormats;
        public List<Hyperlink> Hyperlinks;
        public List<int> LayerMembership;
        public CustomProps CustomProps;
        public VA.VDX.Sections.Geom Geom;

        internal int _id;
        private readonly bool _isGroup;
        private readonly Text m_text = new Text();

        public string Name { get; set; }
        public int Master { get; set; }
        public Page Page;

        private Shape()
        {
            this._id = -1;
            this.Name = null;
        }

        public Shape(int master, double pinx, double piny) :
            this (master, false, pinx, piny)
        {
        }

        public Shape(int master, bool isGroup, double pinx, double piny) :
            this()
        {
            this.Master = master;
            this._isGroup = isGroup;
            this.XForm.PinX.Result = pinx;
            this.XForm.PinY.Result = piny;
        }

        public Shape(int master, double pinx, double piny, double width, double height) :
            this (master, false, pinx, piny, width, height)
        {
        }

        public Shape(int master, bool isGroup, double pinx, double piny, double width, double height) :
            this()
        {
            this.Master = master;
            this._isGroup = isGroup;
            //Get sub shapes

            this.XForm.PinX.Result = pinx;
            this.XForm.PinY.Result = piny;
            this.XForm.Width.Result = width;
            this.XForm.Height.Result = height;
        }

        public Text Text
        {
            get { return this.m_text; }
        }

        public int ID
        {
            get { return _id; }
        }

        public void AddToElement(SXL.XElement parent)
        {
            var shape_el = XMLUtil.CreateVisioSchema2003Element("Shape");
            shape_el.SetAttributeValueInt("ID", this._id);
            shape_el.SetAttributeValue("NameU", this.Name);

            if (this._isGroup)
                shape_el.SetAttributeValue("Type", "Group");
            else
                shape_el.SetAttributeValue("Type", "Shape");


            shape_el.SetAttributeValue("Master", this.Master);

            WriteTransform(shape_el);
            WriteTransform1D(shape_el);
            WriteFill(shape_el);
            WriteLine(shape_el);
            WriteEvent(shape_el);
            WriteLayerMembership(shape_el);
            WriteTextBlock(shape_el);
            WriteProtection(shape_el);
            WriteMisc(shape_el);
            WriteHyperlinks(shape_el);
            WriteTextXForm(shape_el);
            WriteLayout(shape_el);

            WriteCharFormats(shape_el);
            WriteParaFormats(shape_el);
            // TODO: Add support for Tab Stops in VDX
            WriteProps(shape_el);
            WriteGeom(shape_el);
            WriteText(shape_el);


            parent.Add(shape_el);
        }

        private void WriteHyperlinks(XElement shape_el)
        {
            if (this.Hyperlinks == null)
            {
                return;
            }

            int n = 1;
            foreach (var hl in this.Hyperlinks)
            {
                var hyperlink_el = XMLUtil.CreateVisioSchema2003Element("Hyperlink");


                hyperlink_el.SetAttributeValue("NameU", string.Format("Row_{0}",n));
                hyperlink_el.SetAttributeValue("ID", string.Format("{0}", n));

                if (hl.Description != null)
                {
                    var desc_el = XMLUtil.CreateVisioSchema2003Element("Description");
                    desc_el.Value = hl.Description;
                    hyperlink_el.Add(desc_el);
                }

                if (hl.Address != null)
                {
                    var address_el = XMLUtil.CreateVisioSchema2003Element("Address");
                    address_el.Value = hl.Address;
                    hyperlink_el.Add(address_el);
                }

                if (hl.SubAddress != null)
                {
                    var subaddress_el = XMLUtil.CreateVisioSchema2003Element("SubAddress");
                    subaddress_el.Value = hl.SubAddress;
                    hyperlink_el.Add(subaddress_el);
                }

                
                shape_el.Add(hyperlink_el);

                n++;
            }
        }

        private void WriteLayout(XElement xshape)
        {
            if (this.Layout != null)
            {
                this.Layout.AddToElement(xshape);
            }
        }

        private void WriteGeom(XElement xshape)
        {
            if (this.Geom != null)
            {
                this.Geom.AddToElement(xshape, 0);
            }
        }

        private void WriteTransform1D(XElement xshape)
        {
            if (this.XForm1D != null)
            {
                this.XForm1D.AddToElement(xshape);
            }
        }

        private void WriteTextBlock(XElement xshape)
        {
            if (this.TextBlock != null)
            {
                this.TextBlock.AddToElement(xshape);
            }
        }

        private void WriteProtection(XElement xshape)
        {
            if (this.Protection != null)
            {
                this.Protection.AddToElement(xshape);
            }
        }

        private void WriteProps(XElement xshape)
        {
            if (this.CustomProps != null && this.CustomProps.Count > 0)
            {
                foreach (var cp in this.CustomProps)
                {
                    cp.AddToElement(xshape);
                }
            }
        }

        private void WriteParaFormats(XElement xshape)
        {
            if (this.ParaFormats != null)
            {
                int ix = 0;
                foreach (var cf in this.ParaFormats)
                {
                    cf.AddToElement(xshape, ix);
                    ix++;
                }
            }
        }

        private void WriteCharFormats(XElement xshape)
        {
            if (this.CharFormats != null)
            {
                int ix = 0;
                foreach (var cf in this.CharFormats)
                {
                    cf.AddToElement(xshape, ix);
                    ix++;
                }
            }
        }

        private void WriteEvent(XElement xshape)
        {
            if (this.Event != null)
            {
                this.Event.AddToElement(xshape);
            }
        }

        private void WriteTextXForm(XElement xshape)
        {
            if (this.TextXForm != null)
            {
                this.TextXForm.AddToElement(xshape);
            }
        }

        private void WriteMisc(XElement xshape)
        {
            if (this.Misc != null)
            {
                this.Misc.AddToElement(xshape);
            }
        }

        private void WriteTransform(XElement xshape)
        {
            this.XForm.AddToElement(xshape);
        }

        private void WriteFill(XElement xshape)
        {
            if (this.Fill != null)
            {
                this.Fill.AddToElement(xshape);
            }
        }

        private void WriteLine(XElement xshape)
        {
            if (this.Line != null)
            {
                this.Line.AddToElement(xshape);
            }
        }


        private void WriteText(XElement xshape)
        {
            this.Text.AddToElement(xshape);
        }

        private void WriteLayerMembership(XElement xshape)
        {
            if (this.LayerMembership != null)
            {
                if (this.LayerMembership.Count > 0)
                {
                    var xlayermem = XMLUtil.CreateVisioSchema2003Element("LayerMem");
                    var xlayermember = XMLUtil.CreateVisioSchema2003Element("LayerMember");
                    xlayermember.Value = string.Join(";", this.LayerMembership.Select(i => i.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToArray());
                    xlayermem.Add(xlayermember);
                    xshape.Add(xlayermem);
                }
            }
        }

        public static Shape CreateDynamicConnector(Drawing doc)
        {
            int dynamic_connector_id = doc.GetMasterMetaData("Dynamic Connector").ID;
            var shape_el = new Shape(dynamic_connector_id , false, 0, 0);
            shape_el.XForm1D = new VA.VDX.Sections.XForm1D();
            shape_el.XForm1D.BeginX.Formula = "_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)";
            shape_el.XForm1D.BeginX.Result = 0;
            shape_el.XForm1D.BeginY.Formula = "_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)";
            shape_el.XForm1D.BeginY.Result = 0;

            shape_el.XForm1D.EndX.Formula = "_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)";
            shape_el.XForm1D.EndX.Result = 0;
            shape_el.XForm1D.EndY.Formula = "_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)";
            return shape_el;
        }
    }
}