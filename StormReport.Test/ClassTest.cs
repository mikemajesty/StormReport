using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StormReport.Test
{
    public class ClassTest
    {
        private List<ClassTest> List { get; set; }

        public ClassTest()
        {
            List = new List<ClassTest>();
        }

        [ExportableName("Nome do Cliente")]
        public string Nome { get; set; }

        [ExportableName("Idade do Cliente")]
        public int Idade { get; set; }

        [ExportableName("Cidade do Cliente")]
        public string Cidade { get; set; }

        [ExportableName("Estado do Cliente")]
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
