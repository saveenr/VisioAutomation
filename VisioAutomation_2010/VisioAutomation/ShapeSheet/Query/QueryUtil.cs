namespace VisioAutomation.ShapeSheet.Query
{
    class QueryUtil
    {
        internal static CellData[] _combine_formulas_and_results(string[] formulas, string[] results)
        {
            int n = results.Length;

            if (formulas.Length != results.Length)
            {
                throw new System.ArgumentException("Array Lengths must match");
            }

            var combined_data = new ShapeSheet.CellData[n];
            for (int i = 0; i < n; i++)
            {
                combined_data[i] = new ShapeSheet.CellData(formulas[i], results[i]);
            }
            return combined_data;
        }
    }
}