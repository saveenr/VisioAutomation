using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerTools2010
{
    public class SimpleHTML5Writer
    {
        protected System.Xml.XmlWriter _xw;
        private Stack<string> stack;

        protected System.Xml.XmlWriter xmlwriter
        {
            get { return this._xw; }
        }

        public SimpleHTML5Writer(System.Xml.XmlWriter xmlwriter)
        {
            this.stack = new Stack<string>();
            if (xmlwriter == null)
            {
                throw new System.ArgumentNullException("xmlwriter");
            }

            this._xw = xmlwriter;
        }

        public SimpleHTML5Writer(string filename)
        {
            this.stack = new Stack<string>();
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            var settings = new System.Xml.XmlWriterSettings();
            settings.Indent = true;
            settings.OmitXmlDeclaration = true;
            this._xw = System.Xml.XmlWriter.Create(filename, settings);
        }

        public void DocType(string s)
        {
            this.xmlwriter.WriteDocType(s,null,null,null);
        }
        
        public void Start(string s)
        {
            this.xmlwriter.WriteStartElement(s);
            this.stack.Push(s);
        }

        public void End(string s)
        {
            if (stack.Count < 1)
            {
                string msg = string.Format("No matching starting element for <{0}>", s);
                throw new System.ArgumentException(msg, "s");
            }

            string ontop = stack.Pop();
            if (ontop != s)
            {
                string msg = string.Format("Cannot end element <{0}>, expected to end <{1}>", s, ontop);
                throw new System.ArgumentException(msg);
            }

            this.xmlwriter.WriteEndElement();
        }

        public void Element(string name, string s)
        {
            this.xmlwriter.WriteElementString(name, s);
        }

        public void Attribute(string name, string s)
        {
            this.xmlwriter.WriteAttributeString(name, s);
        }

        public void Text(string s)
        {
            this.xmlwriter.WriteString(s);
        }

        public void AttributeIfNotNull(string name, string s)
        {
            if (s != null)
            {
                this.xmlwriter.WriteAttributeString(name, s);
            }
        }

        public void Close()
        {
            this._xw.Close();
        }
    }
}