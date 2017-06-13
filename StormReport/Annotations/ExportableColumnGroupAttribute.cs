using System;

namespace StormReport.Test
{
    public class ExportableColumnGroupAttribute : Attribute
    {
        public string Description { get; set; }
        public ExportableColumnGroupAttribute(string description)
        {
            this.Description = description;
        }
    }
}