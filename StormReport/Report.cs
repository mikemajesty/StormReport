using System.Collections.Generic;
using System.IO;
using System.Web.UI.WebControls;
using System.Drawing;
using System.Web.UI;
using System;
using System.Reflection;
using System.Linq;
using System.Web;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.Diagnostics;
using System.Collections;

namespace StormReport
{
    public class Report
    {
        public void CreateExcelBase<T>(IList<T> listItems, HttpResponseBase Response = null)
        {
            var properties = typeof(T).GetProperties().Where(f => ((ExportableColumnNameAttribute)f.GetCustomAttributes(typeof(ExportableColumnNameAttribute), true).FirstOrDefault()) != null);

            var list = new List<Table>();

            StringBuilder builder = new StringBuilder();
            builder.Append("<table>");
            builder.Append("<tr>");
            builder.Append("</tr>");

            HtmlTable ht = new HtmlTable();

            foreach (var row in listItems.Select(o => new { Properties = properties.Select(g => g), Value = o } ).ToList())
            {
                Debug.WriteLine("<tr>");
                foreach (PropertyInfo cell in row.Properties)
                {
                    var name = ((ExportableColumnNameAttribute)cell.GetCustomAttributes(typeof(ExportableColumnNameAttribute), false).FirstOrDefault()).Description;
                    var cellValue = row.Properties.Select(g => cell.GetValue(row.Value)).FirstOrDefault();
                    Debug.WriteLine("<td>");
                    Debug.WriteLine(cellValue);
                    Debug.WriteLine("</td>");
                }
                Debug.WriteLine("</tr>");
            }

            /*HtmlTable ht = new HtmlTable();
            HtmlTableRow htColumnsRow = new HtmlTableRow();
            HtmlTableCell htCell = new HtmlTableCell();
            htCell.InnerText = prop.Name;
            htColumnsRow.Cells.Add(cell))
            ht.Rows.Add(htColumnsRow);*/

            if (Response == null)
            {
                throw new ArgumentNullException("Response is Required");
            }

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=ClientePerfil.xls");
            Response.ContentType = "application/ms-excel";

            Response.Charset = "utf-8";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);

            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();
        }

        public class Table
        {
            public string Header { get; set; }

            public string Value { get; set; }
        }

        private static Type GetType(PropertyInfo p)
        {
            if (p.PropertyType.IsGenericType && p.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                return (p.PropertyType.GetGenericArguments()[0]).GetType();
            }
            return p.PropertyType;
        }

        public static string Gettable()
        {
            var table = @"<div>
	                        <table cellspacing='0' rules='all' border='1'>
		                        <tr style='background-color:#5F88A4;'>
			                        <th scope='col'>Nome do Cliente</th><th scope='col'>Idade do Cliente</th><th scope='col'>Cidade do Cliente</th><th scope='col'>Estado do Cliente</th>
		                        </tr><tr>
			                        <td>Mike Lima</td><td>28</td><td>Sorocaba</td><td>SP</td>
		                        </tr><tr>
			                        <td>Mike Lima</td><td>28</td><td>Sorocaba</td><td>SP</td>
		                        </tr>
	                        </table>
                        </div>";
            return table;
        }

        public static string GetTableStructure(List<Table> tableList)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("<table>");

            builder.Append("<tr>");

            builder.Append("<th>");

            builder.Append("</th>");

            builder.Append("</tr>");

            builder.Append("<tr>");

            tableList.ForEach(c =>
            {
                builder.Append("<td>");
                builder.Append(c.Value);
                builder.Append("</td>");
            });

            builder.Append("<tr>");

            builder.Append("</table>");

            return builder.ToString();
        }

        public HtmlTable BuildTable<T>(List<T> Data)
        {
            HtmlTable ht = new HtmlTable();
            //Get the columns
            HtmlTableRow htColumnsRow = new HtmlTableRow();
            typeof(T).GetProperties().Select(prop =>
            {
                HtmlTableCell htCell = new HtmlTableCell();
                htCell.InnerText = prop.Name;
                return htCell;
            }).ToList().ForEach(cell => htColumnsRow.Cells.Add(cell));
            ht.Rows.Add(htColumnsRow);
            //Get the remaining rows
            Data.ForEach(delegate (T obj)
            {
                HtmlTableRow htRow = new HtmlTableRow();
                obj.GetType().GetProperties().ToList().ForEach(delegate (PropertyInfo prop)
                {
                    HtmlTableCell htCell = new HtmlTableCell();
                    htCell.InnerText = prop.GetValue(obj, null).ToString();
                    htRow.Cells.Add(htCell);
                });
                ht.Rows.Add(htRow);
            });
            return ht;
        }
    }
}
