using System.Collections.Generic;
using System.Web.UI.WebControls;
using System;
using System.Reflection;
using System.Linq;
using System.Web;
using System.Data;
using System.Text;
using System.IO;
using System.Web.UI;

namespace StormReport
{
    public class TestWithreport
    {
        public string ExcelName { private get; set; }
        public string ExcelTitle { private get; set; }

        public TestWithreport(string excelTitle, string excelName = "ExcelReport")
        {
            this.ExcelName = excelName;
            this.ExcelTitle = excelTitle;
        }

        public void CreateExcelBase<T>(IList<T> listItems, HttpResponseBase Response)
        {
            if (listItems == null)
            {
                throw new ArgumentNullException("Excel list is Required.");
            }

            if (Response == null)
            {
                throw new ArgumentNullException("Response is Required.");
            }

            var properties = GetObjectPropertyInfo<T>();

            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    Table tb = new Table();

                    AddTableAttibute(tb);

                    AddTableTitle<T>(properties, tb);

                    AddTableGroup(properties, tb);

                    AddTableHeader(properties, tb);

                    AddColumnContent(properties, tb, listItems);

                    tb.RenderControl(htw);

                    AddResponseHeader(Response);

                    DownloadExcel(Response, sw);
                }
            }
        }

        private void AddTableAttibute(Table tb)
        {
            tb.Attributes.Add("cellspacing", "0");
            tb.Attributes.Add("rules", "all");
            tb.Attributes.Add("border", "1");
        }

        private void AddColumnContent<T>(IEnumerable<PropertyInfo> properties, Table tb, IList<T> listItems)
        {
            foreach (var row in listItems.Select(o => new { Properties = properties.Select(g => g), Value = o }).ToList())
            {
                TableRow rows = new TableRow();
                rows.TableSection = TableRowSection.TableBody;
                foreach (PropertyInfo cell in row.Properties)
                {
                    var cellValue = row.Properties.Select(g => cell.GetValue(row.Value)).FirstOrDefault();
                    var styleProperty = ((ExportableColumnContentStyleAttribute)cell.GetCustomAttributes(typeof(ExportableColumnContentStyleAttribute), false).FirstOrDefault());
                    TableCell ccell = new TableCell();
                    ccell.Text = cellValue.ToString();
                    rows.Cells.Add(ccell);
                    StringBuilder styles = new StringBuilder();

                    Array.ForEach(styleProperty.Styles, s =>
                    {
                        styles.Append(s.Contains(";") ? s : s + ";");
                    });
                    ccell.Attributes.Add("style", styles.ToString());
                }
                tb.Rows.Add(rows);
            }
        }

        private void AddTableHeader(IEnumerable<PropertyInfo> properties, Table tb)
        {
            TableRow rows = new TableRow();
            rows.TableSection = TableRowSection.TableHeader;
            foreach (var headerCell in properties)
            {
                var headerText = ((ExportableColumnHeaderNameAttribute)headerCell.GetCustomAttributes(typeof(ExportableColumnHeaderNameAttribute), false).FirstOrDefault()).Description;
                var styleProperty = ((ExportableColumnHeaderStyleAttribute)headerCell.GetCustomAttributes(typeof(ExportableColumnHeaderStyleAttribute), false).FirstOrDefault());

                TableHeaderCell hcell = new TableHeaderCell();
                hcell.Text = headerText;
                rows.Cells.Add(hcell);
                StringBuilder styles = new StringBuilder();

                Array.ForEach(styleProperty.Styles, s =>
                {
                    styles.Append(s.Contains(";") ? s : s + ";");
                });

                hcell.Attributes.Add("style", styles.ToString());
            }
            tb.Rows.Add(rows);
        }

        private void AddTableGroup(IEnumerable<PropertyInfo> properties, Table tb)
        {
            TableRow rows = new TableRow();
            rows.TableSection = TableRowSection.TableHeader;
            var columnGroup = properties.Select(c => c.GetCustomAttributes(typeof(ExportableColumnGroupAttribute), false).FirstOrDefault()).ToList();

            foreach (var prop in columnGroup.GroupBy(c => ((ExportableColumnGroupAttribute)c) == null ? null : ((ExportableColumnGroupAttribute)c).Description))
            {
                TableHeaderCell gcell = new TableHeaderCell();
                gcell.Text = prop.Key;
                gcell.ColumnSpan = prop.Count();
                rows.Cells.Add(gcell);
            }
            tb.Rows.Add(rows);
        }

        private void AddTableTitle<T>(IEnumerable<PropertyInfo> properties, Table tb)
        {
            TableRow rows = new TableRow();
            rows.TableSection = TableRowSection.TableHeader;
            StringBuilder style = new StringBuilder();

            Array.ForEach(this.GetTitleStyles<T>(), s =>
            {
                style.Append(s.Contains(";") ? s : s + ";");
            });

            TableHeaderCell hcells = new TableHeaderCell();
            hcells.Text = this.ExcelTitle;
            hcells.ColumnSpan = properties.Count();
            rows.Cells.Add(hcells);
            hcells.Attributes.Add("style", style.ToString());
            tb.Rows.Add(rows);
        }

        private void DownloadExcel(HttpResponseBase Response, StringWriter table)
        {
            Response.Output.Write(table.ToString());
            Response.Flush();
            Response.End();
        }

        private void AddResponseHeader(HttpResponseBase Response)
        {
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=" + GetExcelName());
            Response.ContentType = "application/vnd.ms-excel";
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
