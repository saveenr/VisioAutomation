using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Internal
{
    internal class ShapesheetHelpers
    {
        public static void EnforceValid_ResultType(System.Type result_type)
        {
            if (!IsValidResultType(result_type))
            {
                string msg = string.Format("Unsupported Result Type: {0}", result_type.Name);
                throw new Exceptions.InternalAssertionException(msg);
            }
        }

        public static bool IsValidResultType(System.Type result_type)
        {
            return (result_type == typeof(int)
                    || result_type == typeof(double)
                    || result_type == typeof(string));
        }

        public static void ValidateStreamLengthFormulas(ShapeSheet.Streams.StreamArray stream, object[] formulas)
        {
            if (formulas.Length != stream.Count)
            {
                string msg =
                    string.Format("stream contains {0} items ({1} short values) and requires {2} formula values",
                        stream.Count, stream.Array.Length, stream.Count);
                throw new System.ArgumentException(msg);
            }
        }

        public static void ValidateStreamLengthResults(ShapeSheet.Streams.StreamArray stream, object[] results)
        {
            if (results.Length != stream.Count)
            {
                string msg =
                    string.Format("stream contains {0} items ({1} short values) and requires {2} result values",
                        stream.Count, stream.Array.Length, stream.Count);
                throw new System.ArgumentException(msg);
            }
        }


        public static T[] SystemArrayToTypedArray<T>(System.Array results_sa)
        {
            var results = new T[results_sa.Length];
            results_sa.CopyTo(results, 0);
            return results;
        }

        public static IVisio.VisGetSetArgs GetVisGetSetArgsFromType(System.Type type)
        {
            IVisio.VisGetSetArgs flags;

            if (type == typeof(int))
            {
                flags = IVisio.VisGetSetArgs.visGetTruncatedInts;
            }
            else if (type == typeof(double))
            {
                flags = IVisio.VisGetSetArgs.visGetFloats;
            }
            else if (type == typeof(string))
            {
                flags = IVisio.VisGetSetArgs.visGetStrings;
            }
            else
            {
                string msg = string.Format("Unsupported Result Type: {0}", type.Name);
                throw new Exceptions.InternalAssertionException(msg);
            }

            return flags;
        }
    }
}
