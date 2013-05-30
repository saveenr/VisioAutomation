using System.Collections.Generic;
using System.Linq;
using VisioAutomation.VDX.Internal.Extensions;
using VA=VisioAutomation;
using SXL = System.Xml.Linq;


namespace VisioAutomation.VDX.Elements
{
    public class Drawing : Node
    {
        private readonly PageList _pages;
        private readonly FaceList _faces;
        private readonly List<Window> _windows;
        private readonly List<ColorEntry> _colors;

        private readonly Dictionary<string, MasterMetdata> master_metadata =
            new Dictionary<string, MasterMetdata>(System.StringComparer.OrdinalIgnoreCase);

        public VA.VDX.Sections.DocumentProperties DocumentProperties = new VA.VDX.Sections.DocumentProperties();

        internal int CurrentShapeID = -100;

        public Drawing(SXL.XDocument dom)
        {
            this._pages = new PageList(this);
            this._faces = new FaceList();
            this._windows = new List<Window>();
            this._colors = new List<ColorEntry>();

            var masters_el = dom.Root.ElementVisioSchema2003("Masters");
            if (masters_el == null)
            {
                throw new System.InvalidOperationException();
            }

            // Store information about each master found in the drawing
            foreach (var master_el in masters_el.ElementsVisioSchema2003("Master"))
            {
                var name = master_el.Attribute("NameU").Value;
                var id = int.Parse(master_el.Attribute("ID").Value);

                var subshapes = master_el.Descendants()
                    .Where(el => el.Name.LocalName == "Shape");

                var count_groups = subshapes.Count(shape_el => shape_el.Attribute("Type").Value == "Group");

                bool master_is_group = count_groups > 0;

                var md = new MasterMetdata();
                md.Name = name;
                md.ID = id;
                md.IsGroup = master_is_group;
                md.SubShapeCount = subshapes.Count();

                this.master_metadata[md.Name] = md;

                this.CurrentShapeID = 1;
            }

            var facenames_el = dom.Root.ElementVisioSchema2003("FaceNames");
            foreach (var face_el in facenames_el.ElementsVisioSchema2003("FaceName"))
            {
                var id = int.Parse(face_el.Attribute("ID").Value);
                var name = face_el.Attribute("Name").Value;
                var face = new Face(id, name);
                this._faces.Add(face);
            }

            var colors_el = dom.Root.ElementVisioSchema2003("Colors");
            foreach (var color_el in colors_el.ElementsVisioSchema2003("ColorEntry"))
            {
                var rgb_s = color_el.Attribute("RGB").Value;
                int rgb = int.Parse(rgb_s.Substring(1), System.Globalization.NumberStyles.AllowHexSpecifier);
                var ce = new ColorEntry();
                ce.RGB = rgb;
                this._colors.Add(ce);
            }
        }

        public NamedNodeList<Page> Pages
        {
            get { return _pages; }
        }

        public NamedNodeList<Face> Faces
        {
            get { return _faces; }
        }

        public List<Window> Windows
        {
            get { return _windows; }
        }

        public List<ColorEntry> Colors
        {
            get { return _colors; }
        }

        internal int GetNextShapeID()
        {
            int id = this.CurrentShapeID;
            this.CurrentShapeID++;
            return id;
        }

        public MasterMetdata GetMasterMetData(int id)
        {
            foreach (var m in this.master_metadata)
            {
                if (m.Value.ID == id)
                {
                    return m.Value;
                }
            }

            throw new System.ArgumentException("no such master id");
        }

        public MasterMetdata GetMasterMetaData(string name)
        {
            return this.master_metadata[name];
        }

        public Face AddFace(string name)
        {
            if (!this._faces.ContainsName(name))
            {
                var new_face = new Face(this._faces.Count + 1, name);
                this._faces.Add(new_face);
                return new_face;
            }
            else
            {
                return this._faces[name];
            }
        }

        public static string DefaultTemplateXML
        {
            get { return VA.VDX.Properties.Resources.DefaultVDXTemplate; }
        }

        internal void AccountForMasteSubshapes(int n)
        {
            this.CurrentShapeID += n + 1;
        }
    }
}