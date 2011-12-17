using System.Collections.Generic;

namespace VisioAutomation.Internal
{
    internal class FormatStringParser
    {
        private List<FormatStringSegment> segments;

        public IList<FormatStringSegment> Segments
        {
            get
            {
                return this.segments;
            }
        }

        public FormatStringParser(string format)
        {
            if (format == null)
            {
                throw new System.ArgumentNullException("format");
            }

            this.Parse(format);
        }

        public void Parse(string format)
        {
            if (format == null)
            {
                throw new System.ArgumentNullException("format");
            }

            var pat = new System.Text.RegularExpressions.Regex(@"\{[0-9][0-9]*\}");
            var matches = pat.Matches(format);

            this.segments = new List<FormatStringSegment>();
            foreach (System.Text.RegularExpressions.Match m in matches)
            {
                string captured_string = m.ToString();
                string index_string = captured_string.Substring(1, captured_string.Length - 2);
                int index = System.Int32.Parse(index_string);

                if (index < 0)
                {
                    throw new System.ArgumentException("negative index found");
                }

                var f = new FormatStringSegment(m.ToString(), index, m.Index, m.Index + m.Length);
                this.segments.Add(f);
            }
        }
    }
}