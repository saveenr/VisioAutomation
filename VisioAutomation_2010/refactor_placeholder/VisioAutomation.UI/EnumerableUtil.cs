namespace VisioAutomation.UI
{
    internal static class EnumerableUtil
    {
        /// <summary>
        /// Given a range (start,end) and a number of steps, will yield that a number for each step
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="steps"></param>
        /// <returns></returns>
        public static System.Collections.Generic.IEnumerable<double> RangeSteps(double start, double end, int steps)
        {
            // for non-positive number of steps, yield no points
            if (steps < 1)
            {
                yield break;
            }

            // for exactly 1 step, yield the start value
            if (steps == 1)
            {
                yield return start;
                yield break;
            }

            // for exactly 2 stesp, yield the start value, and then the end value
            if (steps == 2)
            {
                yield return start;
                yield return end;
                yield break;
            }

            // for 3 steps or above, start yielding the segments
            // notice that the start and end values are explicitly returned so that there
            // is no possibility of rounding error affecting their values
            int segments = steps - 1;
            double total_length = end - start;
            double stepsize = total_length / segments;
            yield return start;
            for (int i = 1; i < (steps - 1); i++)
            {
                double p = start + (stepsize * i);
                yield return p;
            }
            yield return end;
        }

        public static void FillArray<T>(T[] array, System.Collections.Generic.IEnumerable<T> items)
        {
            if (array == null)
            {
                throw new System.ArgumentNullException("array");
            }

            if (items == null)
            {
                throw new System.ArgumentNullException("items");
            }

            EnumerableUtil._FillArray(array, items, () => { throw new System.ArgumentException("Not enough items to fill array", "items"); });
        }

        /// <summary>
        /// Places elements from an enumerable into an array. If there are not enough items to fill the array, the default value is used
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="array"></param>
        /// <param name="items"></param>
        /// <param name="default_value"></param>
        public static void FillArray<T>(T[] array, System.Collections.Generic.IEnumerable<T> items, T default_value)
        {
            if (array == null)
            {
                throw new System.ArgumentNullException("array");
            }

            if (items == null)
            {
                throw new System.ArgumentNullException("items");
            }

            EnumerableUtil._FillArray(array, items, () => default_value);
        }

        private static void _FillArray<T>(T[] array, System.Collections.Generic.IEnumerable<T> items, System.Func<T> func_default)
        {
            if (array == null)
            {
                throw new System.ArgumentNullException("array");
            }

            if (items == null)
            {
                throw new System.ArgumentNullException("items");
            }

            if (func_default == null)
            {
                throw new System.ArgumentNullException("func_default");
            }

            using (var e = items.GetEnumerator())
            {
                for (int i = 0; i < array.Length; i++)
                {
                    bool move_ok = e.MoveNext();
                    if (move_ok)
                    {
                        array[i] = e.Current;
                    }
                    else
                    {
                        array[i] = func_default();
                    }
                }
            }
        }
    }
}