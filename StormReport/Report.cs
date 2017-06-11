using System.Collections.Generic;
using System.Web.UI.WebControls;
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
        public void CreateExcelBase<T>(IList<T> listItems, HttpResponseBase Response)
        {
            var properties = GetObjectPropertyInfo<T>();

            var propExcelName = GetExcelFileName<T>();

            if (propExcelName == null)
            {
                throw new ArgumentNullException("Annotation ExcelFileName is Required.");
            }

            var excelName = ((ExcelFileNameAttribute)propExcelName).Name;

            var list = new List<Table>();

            StringBuilder builder = new StringBuilder();
            builder.Append("<table cellspacing='0' rules='all' border='1'>\n");
            builder.Append("<tr>\n");
            foreach (var headerCell in properties)
            {
                builder.Append("<th scope='col'>\n");
                var headerText = ((ExportableColumnNameAttribute)headerCell.GetCustomAttributes(typeof(ExportableColumnNameAttribute), false).FirstOrDefault()).Description;
                builder.Append(headerText);
                builder.Append("</th>\n");
            }
            builder.Append("</tr>\n");

            if (listItems == null)
            {
                throw new ArgumentNullException("Excel list is Required.");
            }

            foreach (var row in listItems.Select(o => new { Properties = properties.Select(g => g), Value = o }).ToList())
            {
                builder.Append("<tr>\n");
                foreach (PropertyInfo cell in row.Properties)
                {
                    var cellValue = row.Properties.Select(g => cell.GetValue(row.Value)).FirstOrDefault();
                    builder.Append("<td scope='row'>\n");
                    builder.Append(cellValue);
                    builder.Append("</td>\n");
                }
                builder.Append("</tr>\n");
            }

            builder.Append("</table>");

            if (Response == null)
            {
                throw new ArgumentNullException("Response is Required.");
            }

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=" + excelName);
            Response.ContentType = "application/ms-excel";

            Response.Charset = "utf-8";

            Response.Output.Write(builder.ToString());
            Response.Flush();
            Response.End();
        }

        private IEnumerable<PropertyInfo> GetObjectPropertyInfo<T>()
        {
            return typeof(T).GetProperties().Where(f => ((ExportableColumnNameAttribute)f.GetCustomAttributes(typeof(ExportableColumnNameAttribute), true).FirstOrDefault()) != null);
        }

        private object GetExcelFileName<T>()
        {
            return typeof(T).GetCustomAttributes(typeof(ExcelFileNameAttribute), true).FirstOrDefault();
        }

        private static Type GetType(PropertyInfo p)
        {
            if (p.PropertyType.IsGenericType && p.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                return (p.PropertyType.GetGenericArguments()[0]).GetType();
            }
            return p.PropertyType;
        }

        public static string GetTable()
        {
            var table = @"<div>
	                        <table cellspacing='0' rules='all' border='1'>
		                        <tr style='background-color:#5F88A4;'>
			                        <th scope='col'>Nome do Cliente</th>
                                    <th scope='col'>Idade do Cliente</th>
                                    <th scope='col'>Cidade do Cliente</th>
                                    <th scope='col'>Estado do Cliente</th>
		                        </tr><tr>
			                        <td>Mike Lima</td>
                                    <td style=mso-number-format:'0\.000'>2800</td>
                                    <td>Sorocaba</td>
                                    <td>SP</td>
                                </tr><tr>
			                        <td>Mike Lima</td>
                                    <td style=mso-number-format:'mm\\/dd\\/yyyy'>06-10-2017 00-00-00</td>
                                    <td style=mso-number-format:'General'>Sorocaba</td>
                                    <td>SP</td>
		                        </tr>
	                        </table>
                        </div>";
            return table;
        }
    }
}
