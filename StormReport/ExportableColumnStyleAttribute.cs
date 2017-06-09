using System;
using System.Collections.Generic;
using System.Web.UI.WebControls;

namespace StormReport
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportableColumnStyleAttribute : Attribute
    {
        public string[] Styles { get; set; }

        public ExportableColumnStyleAttribute(params string[] styles)
        {
            this.Styles = styles;
        }
    }
}
