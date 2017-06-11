using System.Text;

namespace StormReport.BuildTable
{
    public class HtmlTable
    {
        private StringBuilder table;

        public HtmlTable()
        {
            table = new StringBuilder();
        }

        public void InitTable()
        {
            table.Append("<table cellspacing='0' rules='all' border='1'>\n");
        }

        public void EndTable()
        {
            table.Append("</table>\n");
        }

        public void AddRow()
        {
            table.Append("<tr>\n");
        }

        public void EndRow()
        {
            table.Append("</tr>\n");
        }

        public void AddColumnTextHeader(object text)
        {
            table.Append("<th scope='col'>\n");
            table.Append(text);
            table.Append("</th>\n");
        }

        public void AddColumnText(object text)
        {
            table.Append("<td scope='row'>\n");
            table.Append(text);
            table.Append("</td>\n");
        }

        public string ToHtml()
        {
            return table.ToString();
        }
    }
}
