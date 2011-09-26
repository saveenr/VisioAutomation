using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VA=VisioAutomation;
namespace VisioAutomation.CodeGen
{
    public class CSharpWriter
    {
        private System.Text.StringBuilder sb;
        private int indent;

        public CSharpWriter(System.Text.StringBuilder sb)
        {
            this.sb = sb;
            this.indent = 0;
        }
        public void WriteLine()
        {
            sb.AppendLine();
        }

        public void WriteLine(string s)
        {
            for (int i = 0; i < indent; i++)
            {
                sb.Append("\t");
            }
            sb.AppendLine(s);
        }

        public void WriteLine(string fmt, params object[] tokens)
        {
            this.WriteLine(string.Format(fmt,tokens));
        }

        public void Indent()
        {
            this.indent++;
        }

        public void Dedent()
        {
            if (this.indent < 1)
            {
                throw new VA.AutomationException("too much dedent");
            }

            this.indent--;
        }

        public void StartBlock()
        {
            this.WriteLine("{");
            this.Indent();
        }

        public void EndBlock()
        {
            this.Dedent();
            this.WriteLine("}");
        }
    }
}
