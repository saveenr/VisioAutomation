using System.Collections.Generic;

namespace VisioAutomation.Scripting.Utilities
{
    internal static class StringBuilderExtensions
    {
        public static void AppendJoin(this System.Text.StringBuilder sb, string s, IEnumerable<string> tokens)
        {
            // This works exactly like string.Join - except that it appends the results of the join into
            // a StringBuilder object

            // First, make sure the stringbuilder has enough capacity allocated 
            int num_tokens = 0;
            int tokens_length  = 0;

            // use a foreach to minimize the times we have to go through the enumerable
            foreach (string token in tokens)
            {
                num_tokens++;
                tokens_length += token.Length;
            }

            // figure out how much space is needed for the separators
            int num_seps = (num_tokens > 1) ? num_tokens - 1 : 0;
            int separators_length = s.Length*num_seps;

            int combined_length = tokens_length + separators_length;
            int required_capacity = combined_length + sb.Length;
            sb.EnsureCapacity(required_capacity);

            // Now add the tokens and separators
            int i = 0;
            foreach (string token in tokens)
            {
                if (i > 0)
                {
                    sb.Append(s);
                }
                sb.Append(token);
                i++;
            }
        }
    }
}