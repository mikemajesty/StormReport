using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web;

namespace StormReport
{
    public class StormPdf
    {
        public string ExcelName { private get; set; }
        public string ExcelTitle { private get; set; }

        public StormPdf(string excelTitle = "", string excelName = "PdfReport")
        {
            this.ExcelName = excelName;
            this.ExcelTitle = excelTitle;
        }

        public void GeneratePDF<T>(IList<T> listItems, HttpResponseBase Response)
        {
            AddResponseHeader(Response);

            DownloadPdf(Response, new MemoryStream(GeneratePdfBiteArray("")));
        }

        private static byte[] GeneratePdfBiteArray(string htmlTable)
        {
            byte[] bytes;
            using (var ms = new MemoryStream())
            {
                using (var doc = new iTextSharp.text.Document())
                {
                    using (var writer = PdfWriter.GetInstance(doc, ms))
                    {
                        doc.Open();

                        using (var srHtml = new StringReader(htmlTable))
                        {
                            XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, srHtml);
                        }

                        doc.Close();
                    }
                }

                bytes = ms.ToArray();
            }

            return bytes;
        }

        private static void DownloadPdf(HttpResponseBase Response, MemoryStream mstream)
        {
            mstream.WriteTo(Response.OutputStream);
            Response.End();
        }

        private void AddResponseHeader(HttpResponseBase Response)
        {
            Response.Clear();
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=" + GetPdfName());
            Response.Buffer = true;
        }

        private string GetPdfName()
        {
            return this.ExcelName.Contains(".") ? string.Concat(this.ExcelName.Remove(this.ExcelName.LastIndexOf(".")), ".pdf") : string.Concat(this.ExcelName, ".pdf");
        }
    }
}
