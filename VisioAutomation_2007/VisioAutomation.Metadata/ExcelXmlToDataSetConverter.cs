using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using SXL=System.Xml.Linq;

namespace ExcelUtil
{
    class ExcelXmlToDataSetConverter : BaseExcelXmlReader
    {
        private System.Data.DataSet dataset;
        private string worksheetname;
        private System.Data.DataTable dt;

        public ExcelXmlToDataSetConverter ()
        {
        }

        protected override void StartLoad()
        {
            this.dataset = new System.Data.DataSet();
        }

        protected override void EndLoad()
        {
        }

        protected override void StartWorkSheet(string name)
        {
            this.worksheetname = name;
            this.dt = null;
        }

        protected override void EndWorkSheet()
        {
            this.worksheetname = null;
        }

        protected override void StartTable()
        {
            this.dt = new System.Data.DataTable(worksheetname);

            dataset.Tables.Add(dt);
            dt.BeginLoadData();
        }

        protected override void EndTable()
        {
            this.dt.EndLoadData();
            this.dt = null;
        }

        protected override void DefineColumn(ColumnDefinition cd)
        {
            var col = dt.Columns.Add(cd.Name, cd.Type);
        }

        protected override void AddRow(RowData rowdata)
        {
            dt.Rows.Add(rowdata.Value);
        }

        protected override void StartSchema()
        {
        }

        protected override void EndSchema()
        {
        }

        protected override void StartRows()
        {
        }

        protected override void EndRows()
        {
        }

        public System.Data.DataSet DataSet
        {
            get { return this.dataset; }
        }
    }

}
