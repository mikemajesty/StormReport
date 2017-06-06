using System.Collections.Generic;
using System.IO;
using System.Web.UI.WebControls;
using System.Web.Mvc;
using System.Drawing;
using System.Web.UI;
using System;
using System.Reflection;

namespace StormReport
{
    public class Report : Controller
    {
        public void teste<T>(IList<T> list, List<ColumnConfig> config)
        {
            var excel = new System.Data.DataTable("teste");
            excel.Columns.Add("CPF/CNPJ", typeof(double));
            excel.Columns.Add("Data Cadastro", typeof(string));
            excel.Columns.Add("Descrição", typeof(string));
            excel.Columns.Add("Justificação", typeof(string));

            var filds = typeof(T).GetType().GetFields(BindingFlags.Instance | BindingFlags.NonPublic);

            /*for (int row = 0; row < list.Count; row++)
            {
                DataRow newrow = excel.NewRow();
                newrow[excel.Columns["CPF/CNPJ"]] = list[row].CnpjCpf;
                newrow[excel.Columns["Data Cadastro"]] = list[row].DataCadastro;
                newrow[excel.Columns["Descrição"]] = list[row].Descricao;
                newrow[excel.Columns["Justificação"]] = list[row].Justificacao;
                excel.Rows.Add(newrow);
            }*/

            var grid = new GridView();
            grid.DataSource = excel;
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

        public class ColumnConfig
        {
            public string Name { get; set; }
            public Type Type { get; set; }
            public List<TableItemStyle> Styles { get; set; }
        }
    }
}
