using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using StormReport.BuildTable;
using StormReport.Service;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace StormReport
{
    public class StormPdf
    {
        public string PdfName { private get; set; }
        public string PdfTitle { private get; set; }

        public StormPdf(string pdfTitle = "", string pdfName = "PdfReport")
        {
            this.PdfName = pdfName;
            this.PdfTitle = pdfTitle;
        }

        public void CreatePdf<T>(IList<T> listItems, HttpResponseBase Response)
        {
            if (listItems == null)
                throw new ArgumentNullException("Excel list is Required.");

            if (Response == null)
                throw new ArgumentNullException("Response is Required.");

            var properties = ReportService.GetObjectPropertyInfo<T>();

            TableFactory table = new TableFactory();
            table.InitTable();

            AddPdfTitle<T>(properties.Count(), table);

            ReportService.AddColumnGroup(properties, table);

            ReportService.AddTableColumnHeader(properties, table);

            ReportService.AddTableColumnCell(properties, table, listItems);

            table.EndTable();

            AddResponseHeader(Response);

            DownloadPdf(Response, new MemoryStream(GeneratePdfBiteArray(table.ToHtml())));
        }

        private static void DownloadPdf(HttpResponseBase Response, MemoryStream mstream)
        {
            mstream.WriteTo(Response.OutputStream);
            Response.ClearContent();
            Response.Clear();
            Response.End();
        }

        private void AddResponseHeader(HttpResponseBase Response)
        {
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=" + GetPdfName());
            Response.Buffer = true;
        }

        private string GetPdfName()
        {
            return this.PdfName.Contains(".") ? string.Concat(this.PdfName.Remove(this.PdfName.LastIndexOf(".")), ".pdf") : string.Concat(this.PdfName, ".pdf");
        }

        private void AddPdfTitle<T>(int columnCount, TableFactory table)
        {
            if (!string.IsNullOrEmpty(this.PdfTitle))
            {
                table.AddRow();
                table.AddReportTitle(this.PdfTitle, columnCount, ReportService.GetTitleStyles<T>());
                table.EndRow();
            }
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
    }
}
