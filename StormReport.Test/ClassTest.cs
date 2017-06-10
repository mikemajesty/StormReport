using System.Collections.Generic;

namespace StormReport.Test
{
    [ExcelFileName("ExcelReport")]
    public class ClassTest
    {
        private List<ClassTest> List { get; set; }

        public ClassTest()
        {
            List = new List<ClassTest>();
        }

        [ExportableColumnName("Nome do Cliente")]
        public string Nome { get; set; }

        [ExportableColumnName("Idade do Cliente")]
        public int Idade { get; set; }

        [ExportableColumnName("Cidade do Cliente")]
        public string Cidade { get; set; }

        [ExportableColumnName("Estado do Cliente")]
        public string Estado { get; set; }

        public void Add(ClassTest e)
        {
            List.Add(e);
        }

        public List<ClassTest> GetList()
        {
            return List;
        }
    }
}
