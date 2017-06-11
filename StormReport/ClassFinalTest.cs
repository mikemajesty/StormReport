using System;
using System.Collections.Generic;

namespace StormReport.Test
{
    [ExcelFileName("ExcelReport", useDateTimeToAdditionalName: true)]
    public class ClassFinalTest
    {
        private List<ClassFinalTest> List;

        public ClassFinalTest()
        {
            List = new List<ClassFinalTest>();
        }
  
        [ExportableColumnName("Nome do Cliente")]
        public string Nome { get; set; }

        [ExportableColumnName("Idade do Cliente")]
        public int Idade { get; set; }

        [ExportableColumnName("Cidade do Cliente")]
        public string Cidade { get; set; }

        [ExportableColumnName("Estado do Cliente")]
        public string Estado { get; set; }

        public void Add(ClassFinalTest e)
        {
            List.Add(e);
        }

        public List<ClassFinalTest> GetList()
        {
            return List;
        }
    }
}
