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

namespace StormReport
{
    public class Report
    {
        public void CreateExcelBase<T>(IList<T> listItems, HttpResponseBase Response = null)
        {
            var excelExportable = new DataTable("ExportToExcelTable");
            DataRow newrow = null;
            var properties = typeof(T).GetProperties().Where(f => ((ExportableNameAttribute)f.GetCustomAttributes(typeof(ExportableNameAttribute), true).FirstOrDefault()) != null);
            foreach (PropertyInfo prop in properties)
            {
                var name = ((ExportableNameAttribute)prop.GetCustomAttributes(typeof(ExportableNameAttribute), false).FirstOrDefault()).Description;

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
    }
}
