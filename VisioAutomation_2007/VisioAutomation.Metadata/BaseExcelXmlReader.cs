using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using SXL=System.Xml.Linq;

namespace ExcelUtil
{
    abstract class BaseExcelXmlReader
    {
        private List<ColumnDefinition> _columns;
        private string source_uri;

        protected abstract void StartLoad();
        protected abstract void EndLoad();
        protected abstract void StartWorkSheet(string name);
        protected abstract void EndWorkSheet();
        protected abstract void StartTable();
        protected abstract void EndTable();
        protected abstract void AddRow(RowData row);
        protected abstract void StartSchema();
        protected abstract void EndSchema();
        protected abstract void StartRows();
        protected abstract void EndRows();
        protected abstract void DefineColumn(ColumnDefinition col);

        public BaseExcelXmlReader()
        {
        }

        public List<ColumnDefinition> Columns
        {
            get { return _columns; }
        }

        public void Load(string uri)
        {
            this.source_uri = uri;
            var doc = SXL.XDocument.Load(this.source_uri);
            handle_xml(doc);
        }

        public void Parse(string text)
        {
            this.source_uri = null;
            var doc = SXL.XDocument.Parse(text);
            handle_xml(doc);
        }

        private void handle_xml(XDocument doc)
        {
            var workbook_ns = SXL.XNamespace.Get(@"urn:schemas-microsoft-com:office:spreadsheet");
            var worksheet_els =
                doc.Elements(workbook_ns + "Workbook")
                    .Elements(workbook_ns + "Worksheet")
                    .ToList();

            this.StartLoad();

            foreach (var worksheet_el in worksheet_els)
            {
                string worksheetname = worksheet_el.Attribute(workbook_ns + "Name").Value;

                this.StartWorkSheet(worksheetname);

                foreach (var table_els in worksheet_el.Elements(workbook_ns + "Table"))
                {
                    this._columns = new List<ColumnDefinition>();
                    this.StartTable();

                    var col_els = table_els.Elements(workbook_ns + "Column").ToList();
                    int num_cols = col_els.Count;
                    var rowdata = new RowData(num_cols);

                    int row_index = 0;
                    var row_els = table_els.Elements(workbook_ns + "Row");
                    foreach (var el_row in row_els)
                    {
                        var cell_els = el_row.Elements(workbook_ns + "Cell").ToList();

                        // If the current list of cells is more than the obj_array can handle, reallocate the obj_array
                        if (cell_els.Count > rowdata.Length)
                        {
                            rowdata = new RowData(cell_els.Count);
                        }

                        // clean the row
                        rowdata.Clear();

                        // fill the obj_array with the values from the cells
                        int column_index = 0;
                        foreach (var cell_el in cell_els)
                        {
                            var el_data = cell_el.Element(workbook_ns + "Data");
                            if (el_data != null)
                            {
                                // store the data in the rowdata
                                var el_type = el_data.Attribute(workbook_ns + "Type");
                                string type_str = el_type == null ? null : el_type.Value;
                                rowdata.Value[column_index] = el_data.Value;
                                rowdata.Type[column_index] = type_str;
                            }
                            else
                            {
                                // no data -> do nothing (yes this is a possible case we must handle)
                                rowdata.Value[column_index] = null;
                                rowdata.Type[column_index] = null;
                            }

                            column_index++;
                        }

                        // Handle each Row
                        if (row_index == 0)
                        {
                            this.StartSchema();
                            // If it is the first row then use it for colyumn names
                            for (int i = 0; i < rowdata.Length; i++)
                            {
                                string col_value = rowdata.Value[i];
                                string col_name = col_value != null ? (string) col_value : string.Format("Col{0}", i + 1);

                                var coldef = new ColumnDefinition(col_name, typeof (string));

                                this._columns.Add(coldef);
                                this.DefineColumn(coldef);
                            }
                            this.EndSchema();
                            this.StartRows();
                        }
                        else
                        {
                            // for every other row, just add it to the table
                            this.AddRow(rowdata);
                        }
                        row_index++;
                    }
                    this.EndRows();
                    this.EndTable();
                }

                this.EndWorkSheet();
            }

            this.EndLoad();
        }
    }
}
