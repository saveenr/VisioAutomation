namespace VisioPowerShell.Commands.VisioPageCells;

[SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioPageCells)]
public class GetVisioPageCells: VisioCmdlet
{

    [SMA.Parameter(Mandatory = false)]
    public string[] Cell { get; set; }


    [SMA.Parameter(Mandatory = false)]
    public SMA.SwitchParameter Results { get; set; }

    [SMA.Parameter(Mandatory = false)]
    public Models.ResultType ResultType = VisioPowerShell.Models.ResultType.String;

    // CONTEXT:PAGES
    [SMA.Parameter(Mandatory = false)]
    public IVisio.Page[] Page { get; set; }

    protected override void ProcessRecord()
    {
        var valuetype = this.Results
            ? VisioAutomation.ShapeSheet.CellValueType.Result
            : VisioAutomation.ShapeSheet.CellValueType.Formula;

        var targetpages = new VisioScripting.TargetPages(this.Page).ResolveToPages(this.Client);

        if (targetpages.Pages.Count < 1)
        {
            return;
        }

        var template = new VisioPowerShell.Models.PageCells();
        var dicof_name_to_cell = VisioPowerShell.Internal.NamedSrcDictionary.FromCells(template);
        var desired_cells = this.Cell ?? dicof_name_to_cell.Keys.ToArray();
        var query = _create_query(dicof_name_to_cell, desired_cells);
            
        var datatable = new System.Data.DataTable();

        foreach (var page in targetpages.Pages)
        {
            var shapesheet = page.PageSheet;
            var shapeids = new List<int> { shapesheet.ID };
            var surface = new VisioAutomation.SurfaceTarget(page);
            var temp_datatable = VisioPowerShell.Internal.DataTableHelpers.QueryToDataTable(query, valuetype, this.ResultType, shapeids, surface);
            datatable.Merge(temp_datatable);
        }

        // Annotate the returned datatable to disambiguate rows
        var pageid_col = datatable.Columns.Add("PageID", typeof(int));
        int pageid_colindex = 0;
        pageid_col.SetOrdinal(pageid_colindex);
        foreach (int row_index in Enumerable.Range(0,targetpages.Pages.Count))
        {
            var page = targetpages.Pages[row_index];
            datatable.Rows[row_index][pageid_colindex] = page.ID;
        }

        this.WriteObject(datatable);
    }

    private VisioAutomation.ShapeSheet.Query.CellQuery _create_query(
        VisioPowerShell.Internal.NamedSrcDictionary celldic,
        IList<string> cellnames)
    {
        var invalid_names = cellnames.Where(cellname => !celldic.ContainsKey(cellname)).ToList();

        if (invalid_names.Count > 0)
        {
            var quoted_names = invalid_names.Select( s=> string.Format("\"{0}\"",s));
            string msg = "Invalid cell names: " + string.Join(",", quoted_names);
            throw new System.ArgumentException(nameof(cellnames),msg);
        }

        var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

        foreach (string cellname in cellnames)
        {
            // resolve any wildcards to the actual cell names
            var resolved_cellnames = celldic.ExpandKeyWildcard(cellname);

            foreach (var resolved_cellname in resolved_cellnames)
            {
                if (!query.Columns.Contains(resolved_cellname))
                {
                    var resolved_src = celldic[resolved_cellname];
                    query.Columns.Add(resolved_src, resolved_cellname);
                }
            }
        }

        return query;
    }
}