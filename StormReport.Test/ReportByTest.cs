using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Collections;
using static StormReport.Report;
using System.Web.Mvc;

namespace StormReport.Test
{
    [TestClass]
    public class ReportByTest
    {
        [TestMethod]
        public void TestMethod1()
        {
            Report e = new Report();
            var d = new List<ColumnConfig>();
            var d1 = new ColumnConfig { Name = "CPF/CNPJ", Type= typeof(double) };
            d.Add(d1);
            var d2 = new ColumnConfig { Name = "Data Cadastro", Type = typeof(string) };
            d.Add(d2);
            var d3 = new ColumnConfig { Name = "Descrição", Type = typeof(string) };
            d.Add(d3);
            var d4 = new ColumnConfig { Name = "Justificação", Type = typeof(string) };
            d.Add(d4);

            var r = new test { Cidade = "Ibiúna", Estado = "SP", Idade = 10, Nome = "Mike" };
            var r2 = new test { Cidade = "Sorocaba", Estado = "SP", Idade = 28, Nome = "Mike Lima" };
            var f = new List<test>();
            f.Add(r);
            f.Add(r2);
            e.CreateExcelBase(f);
        }

        public class test
        {
            public List<test> List { get; set; }

            public test()
            {
                List = new List<test>();
            }

            [ExportableName("Nome do Cliente")]
            public string Nome { get; set; }

            [ExportableName("Idade do Cliente")]
            public int Idade { get; set; }

            [ExportableName("Cidade do Cliente")]
            public string Cidade { get; set; }

            [ExportableName("Estado do Cliente")]
            public string Estado { get; set; }

            public void Add(test e)
            {
                List.Add(e);
            }
        }
    }
}
