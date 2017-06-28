using System.Collections.Generic;
using System;
using System.Linq;
using System.Web;
using StormReport.BuildTable;
using System.Text;
using StormReport.Service;

namespace StormReport
{
    public class StormExcel
    {
        public string ExcelName { private get; set; }
        public string ExcelTitle { private get; set; }

        public StormExcel(string excelTitle = "", string excelName = "ExcelReport")
        {
            this.ExcelName = excelName;
            this.ExcelTitle = excelTitle;
        }

        public void CreateExcel<T>(IList<T> listItems, HttpResponseBase Response)
        {
            if (listItems == null)
                throw new ArgumentNullException("Excel list is Required.");

            if (Response == null)
                throw new ArgumentNullException("Response is Required.");

            var properties = ReportService.GetObjectPropertyInfo<T>();

            TableFactory table = new TableFactory();
            table.InitTable();

            AddExcelTitle<T>(properties.Count(), table);

            ReportService.AddColumnGroup(properties, table);

            ReportService.AddTableColumnHeader(properties, table);

            ReportService.AddTableColumnCell(properties, table, listItems);

            table.EndTable();

            AddResponseHeader(Response);

            DownloadExcel(Response, table);
        }

        private void DownloadExcel(HttpResponseBase Response, TableFactory table)
        {
            Response.Output.Write(table.ToHtml());
            Response.Flush();
            Response.Clear();
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

        private void AddExcelTitle<T>(int columnCount, TableFactory table)
        {
            if (!string.IsNullOrEmpty(this.ExcelTitle))
            {
                table.AddRow();
                table.AddReportTitle(this.ExcelTitle, columnCount, ReportService.GetTitleStyles<T>());
                table.EndRow();
            }
        }
    }
}
