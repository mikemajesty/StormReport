using System.Collections.Generic;
using System.Web.UI.WebControls;
using System;
using System.Reflection;
using System.Linq;
using System.Web;
using System.Data;
using StormReport.BuildTable;
using System.Text;

namespace StormReport
{
    public class Report
    {
        public string ExcelName { private get; set; }
        public string ExcelTitle { private get; set; }

        public Report(string excelTitle, string excelName = "ExcelReport")
        {
            this.ExcelName = excelName;
            this.ExcelTitle = excelTitle;
        }

        public void CreateExcelBase<T>(IList<T> listItems, HttpResponseBase Response)
        {
            var properties = GetObjectPropertyInfo<T>();

            var excelName = this.ExcelName;

            var list = new List<Table>();

            HtmlTable table = new HtmlTable();
            table.InitTable();

            AddExcelTitle<T>(table, properties.Count());

            AddColumnGroup(properties, table);

            AddTableColumnHeader(properties, table);

            if (listItems == null)
            {
                throw new ArgumentNullException("Excel list is Required.");
            }

            AddTableColumnCell(listItems, properties, table);

            table.EndTable();

            if (Response == null)
            {
                throw new ArgumentNullException("Response is Required.");
            }

            AddResponseHeader(Response);
            DownloadExcel(Response, table);
        }

        private void AddColumnGroup(IEnumerable<PropertyInfo> properties, HtmlTable table)
        {
            var columnGroup = properties.Select(c => c.GetCustomAttributes(typeof(ExportableColumnGroupAttribute), false).FirstOrDefault()).ToList();

            if (columnGroup.FirstOrDefault(c => ((ExportableColumnGroupAttribute)c) != null) == null)
                return;

            table.AddRow();
            foreach (var prop in columnGroup.GroupBy(c => ((ExportableColumnGroupAttribute)c) == null ? null : ((ExportableColumnGroupAttribute)c).Description))
            {
                table.AddColumnGroup(prop.Key ?? "", prop.Count());
            }
            table.EndRow();
        }

        private void AddExcelTitle<T>(HtmlTable table, int columnCount)
        {
            table.AddRow();
            table.AddExcelTitle(this.ExcelTitle, columnCount, this.GetTitleStyles<T>());
            table.EndRow();
        }

        private static void AddTableColumnCell<T>(IList<T> listItems, IEnumerable<PropertyInfo> properties, HtmlTable table)
        {
            foreach (var row in listItems.Select(o => new { Properties = properties.Select(g => g), Value = o }).ToList())
            {
                table.AddRow();
                foreach (PropertyInfo cell in row.Properties)
                {
                    var cellValue = row.Properties.Select(g => cell.GetValue(row.Value)).FirstOrDefault();
                    var styleProperty = ((ExportableColumnContentStyleAttribute)cell.GetCustomAttributes(typeof(ExportableColumnContentStyleAttribute), false).FirstOrDefault());
                    var addtionalText = ((ExportableAddtionalTextAttribute)cell.GetCustomAttributes(typeof(ExportableAddtionalTextAttribute), false).FirstOrDefault());

                    table.AddColumnText(cellValue, styleProperty.Styles ?? new string[] { }, addtionalText);
                }
                table.EndRow();
            }
        }

        private static void AddTableColumnHeader(IEnumerable<PropertyInfo> properties, HtmlTable table)
        {
            table.AddRow();
            foreach (var headerCell in properties)
            {
                var headerText = ((ExportableColumnHeaderNameAttribute)headerCell.GetCustomAttributes(typeof(ExportableColumnHeaderNameAttribute), false).FirstOrDefault()).Description;
                var styleProperty = ((ExportableColumnHeaderStyleAttribute)headerCell.GetCustomAttributes(typeof(ExportableColumnHeaderStyleAttribute), false).FirstOrDefault());

                table.AddColumnTextHeader(headerText, styleProperty.Styles ?? new string[] { });
            }
            table.EndRow();
        }

        private void DownloadExcel(HttpResponseBase Response, HtmlTable table)
        {
            Response.Output.Write(table.ToHtml());
            Response.Flush();
            Response.End();
        }

        private void AddResponseHeader(HttpResponseBase Response)
        {
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=" + GetExcelName());
            Response.ContentType = "application/ms-excel";
            Response.Charset = Encoding.UTF8.EncodingName;
            Response.ContentType = "application/text";
            Response.ContentEncoding = Encoding.Unicode;
            Response.BinaryWrite(Encoding.Unicode.GetPreamble());
        }

        private string GetExcelName()
        {
            return this.ExcelName.Contains(".") ? string.Concat(this.ExcelName.Remove(this.ExcelName.LastIndexOf(".")), ".xls") : string.Concat(this.ExcelName, ".xls");
        }

        private IEnumerable<PropertyInfo> GetObjectPropertyInfo<T>()
        {
            return typeof(T).GetProperties().Where(f => ((ExportableColumnHeaderNameAttribute)f.GetCustomAttributes(typeof(ExportableColumnHeaderNameAttribute), true).FirstOrDefault()) != null);
        }

        private string[] GetTitleStyles<T>()
        {
            var propExcelName = typeof(T).GetCustomAttributes(typeof(ExcelTitleStyleAttribute), true).FirstOrDefault();
            return ((ExcelTitleStyleAttribute)propExcelName) != null ? ((ExcelTitleStyleAttribute)propExcelName).Styles : new string[] { };
        }
    }
}
