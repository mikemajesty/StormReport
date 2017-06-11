using System;
using System.Collections.Generic;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace StormReport
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportableColumnStyleAttribute : Attribute
    {
        public HtmlTextWriterStyle[] Styles { get; set; }

        public ExportableColumnStyleAttribute(params HtmlTextWriterStyle[] styles)
        {
            this.Styles = styles;
        }
    }
}
