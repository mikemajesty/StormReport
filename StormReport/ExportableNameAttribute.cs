using System;

namespace StormReport
{
    public class ExportableNameAttribute : Attribute
    {
        public string Description { get; set; }

        public ExportableNameAttribute(string desc)
        {
            this.Description = string.IsNullOrEmpty(desc) ? "Unknown name" : desc;
        }
    }
}
