using System;
using System.Collections.Generic;
using System.Web.UI.WebControls;

namespace StormReport
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportableColumnNameAttribute : Attribute
    {
        public string Description { get; set; }
      
        public ExportableColumnNameAttribute(string description)
        {
            this.Description = string.IsNullOrEmpty(description) ? "Unknown name" : description;
        }
    }
}
