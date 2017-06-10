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

namespace StormReport
{
    public class Report
    {
        public void CreateExcelBase<T>(IList<T> listItems, HttpResponseBase Response = null)
        {
            var excelExportable = new DataTable("ExportToExcelTable");
            DataRow newrow = null;
            var properties = typeof(T).GetProperties().Where(f => ((ExportableColumnNameAttribute)f.GetCustomAttributes(typeof(ExportableColumnNameAttribute), true).FirstOrDefault()) != null);
            foreach (PropertyInfo prop in properties)
            {
                var name = ((ExportableColumnNameAttribute)prop.GetCustomAttributes(typeof(ExportableColumnNameAttribute), false).FirstOrDefault()).Description;

                var propertiesValues = listItems.Select(o => prop.GetValue(o)).ToList();
                excelExportable.Columns.Add(name, GetType(prop));

                if (newrow == null)
                {
                    propertiesValues.ForEach(r =>
                    {
                        newrow = excelExportable.NewRow();
                        excelExportable.Rows.Add(newrow);
                    });
                }

                for (int cont = 0; cont < propertiesValues.Count; cont++)
                {
                    newrow[excelExportable.Columns[name]] = propertiesValues[cont];
                    excelExportable.Rows[cont].ItemArray = newrow.ItemArray;
                }
            }

            if (Response == null)
            {
                throw new ArgumentNullException("Response is Required");
            }

            var grid = new GridView();
            grid.DataSource = excelExportable;
            grid.DataBind();
            grid.HeaderStyle.BackColor = Color.FromArgb(95, 136, 164);
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=ClientePerfil.xls");
            Response.ContentType = "application/ms-excel";

            Response.Charset = "utf-8";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);
            grid.RenderControl(htw);

            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();
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

        public static string GetTableStructure()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("<table>");

            builder.Append("<tr>");

            builder.Append("<th>");
            builder.Append("</th>");

            builder.Append("</tr>");

            builder.Append("<tr>");

            builder.Append("<td>");
            builder.Append("</td>");

            builder.Append("<tr>");

            builder.Append("</table>");

            return builder.ToString();
        }
    }
}
