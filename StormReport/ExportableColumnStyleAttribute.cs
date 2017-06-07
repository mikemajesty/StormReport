using System;
using System.Collections.Generic;
using System.Web.UI.WebControls;

namespace StormReport
{
    public class ExportableColumnStyleAttribute : Attribute
    {
        public TableItemStyle Styles { get; set; }

        public ExportableColumnStyleAttribute(TableItemStyle styles)
        {
            this.Styles = styles;
        }
    }
}
