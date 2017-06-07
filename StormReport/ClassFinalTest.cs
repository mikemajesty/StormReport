using System.Collections.Generic;

namespace StormReport.Test
{
    public class ClassFinalTest
    {
        private List<ClassFinalTest> List { get; set; }

        public ClassFinalTest()
        {
            List = new List<ClassFinalTest>();
        }

        [ExportableColumnName("Nome do Cliente")]
        [ExportableColumnStyle(null)]
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
