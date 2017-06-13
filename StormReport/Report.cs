using System.Collections.Generic;
using System.Web.UI.WebControls;
using System;
using System.Reflection;
using System.Linq;
using System.Web;
using System.Data;
using System.Text;
using StormReport.BuildTable;
using StormReport.Test;

namespace StormReport
{
    public class Report
    {
        public void CreateExcelBase<T>(IList<T> listItems, HttpResponseBase Response)
        {
            var properties = GetObjectPropertyInfo<T>();

            var propExcelName = GetExcelFileName<T>();

            if (propExcelName == null)
            {
                throw new ArgumentNullException("Annotation ExcelFileName is Required.");
            }

            var excelName = GetExcelName(propExcelName);

            var list = new List<Table>();

            HtmlTable table = new HtmlTable();
            table.InitTable();

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

            AddResponseHeader(Response, excelName);
            DownloadExcel(Response, table);
        }

        private static string GetExcelName(object propExcelName)
        {
            return ((ExcelFileNameAttribute)propExcelName).Name;
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

        private void AddResponseHeader(HttpResponseBase Response, string excelName)
        {
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=" + excelName);
            Response.ContentType = "application/ms-excel";
            Response.Charset = "utf-8";
            //Response.ContentEncoding = Encoding.Unicode;
            //Response.BinaryWrite(Encoding.Unicode.GetPreamble());
        }

        private IEnumerable<PropertyInfo> GetObjectPropertyInfo<T>()
        {
            return typeof(T).GetProperties().Where(f => ((ExportableColumnHeaderNameAttribute)f.GetCustomAttributes(typeof(ExportableColumnHeaderNameAttribute), true).FirstOrDefault()) != null);
        }

        private object GetExcelFileName<T>()
        {
            return typeof(T).GetCustomAttributes(typeof(ExcelFileNameAttribute), true).FirstOrDefault();
        }
    }
}
