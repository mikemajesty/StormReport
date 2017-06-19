using System.Collections.Generic;
using System;
using System.Reflection;
using System.Linq;
using System.Web;
using System.Data;
using StormReport.BuildTable;
using System.Text;

namespace StormReport
{
    public class StormExcel
    {
        public string ExcelName { private get; set; }
        public string ExcelTitle { private get; set; }

        public StormExcel(string excelTitle, string excelName = "ExcelReport")
        {
            this.ExcelName = excelName;
            this.ExcelTitle = excelTitle;
        }

        public void CreateExcelBase<T>(IList<T> listItems, HttpResponseBase Response)
        {
            if (listItems == null)
                throw new ArgumentNullException("Excel list is Required.");

            if (Response == null)
                throw new ArgumentNullException("Response is Required.");

            var properties = GetObjectPropertyInfo<T>();

            TableFactory table = new TableFactory();
            table.InitTable();

            AddExcelTitle<T>(properties.Count(), table);

            AddColumnGroup(properties, table);

            AddTableColumnHeader(properties, table);

            AddTableColumnCell(properties, table, listItems);

            table.EndTable();

            AddResponseHeader(Response);

            DownloadExcel(Response, table);
        }

        private void AddColumnGroup(IEnumerable<PropertyInfo> properties, TableFactory table)
        {
            var columnGroup = properties.Select(c => c.GetCustomAttributes(typeof(ExportableColumnGroupAttribute), false).FirstOrDefault()).ToList();

            if (columnGroup.FirstOrDefault(c => ((ExportableColumnGroupAttribute)c) != null) == null)
                return;

            table.AddRow();
            foreach (var prop in columnGroup.GroupBy(c => ((ExportableColumnGroupAttribute)c) == null ? null : ((ExportableColumnGroupAttribute)c).Description))
            {
                var style = properties.Select(c => c.GetCustomAttributes(typeof(ExportableColumnGroupAttribute), false)).SelectMany(c => c.Where(p => ((ExportableColumnGroupAttribute)p) != null && ((ExportableColumnGroupAttribute)p).Styles.Count() > 0)).Where(p => ((ExportableColumnGroupAttribute)p).Description == prop.Key);
                table.AddColumnGroup(prop.Key ?? "", prop.Count(), style.FirstOrDefault() != null ? ((ExportableColumnGroupAttribute)style.FirstOrDefault()).Styles : new string[] { });
            }
            table.EndRow();
        }

        private void AddExcelTitle<T>(int columnCount, TableFactory table)
        {
            if (!string.IsNullOrEmpty(this.ExcelTitle))
            {
                table.AddRow();
                table.AddExcelTitle(this.ExcelTitle, columnCount, this.GetTitleStyles<T>());
                table.EndRow();
            }
        }

        private static void AddTableColumnCell<T>(IEnumerable<PropertyInfo> properties, TableFactory table, IList<T> listItems)
        {
            foreach (var row in listItems.Select(o => new { Properties = properties.Select(g => g), Value = o }).ToList())
            {
                table.AddRow();
                foreach (PropertyInfo cell in row.Properties)
                {
                    var cellValue = row.Properties.Select(g => cell.GetValue(row.Value)).FirstOrDefault();
                    var styleProperty = ((ExportableColumnContentStyleAttribute)cell.GetCustomAttributes(typeof(ExportableColumnContentStyleAttribute), false).FirstOrDefault());
                    var addtionalText = ((ExportableAddtionalTextAttribute)cell.GetCustomAttributes(typeof(ExportableAddtionalTextAttribute), false).FirstOrDefault());

                    table.AddColumnContentText(cellValue, styleProperty != null ? styleProperty.Styles : new string[] { }, addtionalText);
                }
                table.EndRow();
            }
        }

        private static void AddTableColumnHeader(IEnumerable<PropertyInfo> properties, TableFactory table)
        {
            table.AddRow();
            foreach (var headerCell in properties)
            {
                var headerText = ((ExportableColumnHeaderNameAttribute)headerCell.GetCustomAttributes(typeof(ExportableColumnHeaderNameAttribute), false).FirstOrDefault()).Description;
                var styleProperty = ((ExportableColumnHeaderStyleAttribute)headerCell.GetCustomAttributes(typeof(ExportableColumnHeaderStyleAttribute), false).FirstOrDefault());

                table.AddColumnTextHeader(headerText, styleProperty  != null ? styleProperty.Styles : new string[] { });
            }
            table.EndRow();
        }

        private void DownloadExcel(HttpResponseBase Response, TableFactory table)
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
