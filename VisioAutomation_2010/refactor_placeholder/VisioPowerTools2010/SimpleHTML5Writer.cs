using System;
using System.Collections.Generic;
using System.Xml;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerTools2010
{
    public class SimpleHtml5Writer : IDisposable
    {
        protected XmlWriter _xw;
        private readonly Stack<string> _element_stack;

        public SimpleHtml5Writer(XmlWriter xmlwriter)
        {
            if (xmlwriter == null)
            {
                throw new ArgumentNullException("xmlwriter");
            }

            this._xw = xmlwriter;
            this._element_stack = new Stack<string>();          
        }

        public SimpleHtml5Writer(string filename)
        {
            if (filename == null)
            {
                throw new ArgumentNullException("filename");
            }
            var settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.OmitXmlDeclaration = true;
            this._xw = XmlWriter.Create(filename, settings);
            this._element_stack = new Stack<string>();
        }

        public void DocType(string s)
        {
            string pubid = null;
            string sysid = null;
            string subset = null;
            this._xw.WriteDocType(s,pubid,sysid,subset);
        }

        public void Start(string s)
        {
            this._xw.WriteStartElement(s);
            this._element_stack.Push(s);
        }

        public void End(string s)
        {
            if (this._element_stack.Count < 1)
            {
                string msg = string.Format("No matching starting element for <{0}>", s);
                throw new ArgumentException(msg, "s");
            }

            string ontop = this._element_stack.Pop();
            if (ontop != s)
            {
                string msg = string.Format("Cannot end element <{0}>, expected to end <{1}>", s, ontop);
                throw new ArgumentException(msg);
            }

            this._xw.WriteEndElement();
        }

        public void Element(string name, string s)
        {
            this._xw.WriteElementString(name, s);
        }

        public void Attribute(string name, string s)
        {
            this._xw.WriteAttributeString(name, s);
        }

        public void Text(string s)
        {
            this._xw.WriteString(s);
        }

        public void AttributeIfNotNull(string name, string s)
        {
            if (s != null)
            {
                this._xw.WriteAttributeString(name, s);
            }
        }

        public void Close()
        {
            this._xw.Close();
        }

        public void Dispose()
        {
            this.Close();
        }
    }
}