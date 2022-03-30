using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxReportNetDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<DemoDto> demoDtos = new List<DemoDto>();
            demoDtos.Add(new DemoDto() { Name = "Stanley", Desription = "測試文件一說明", Title = "測試文件一" });
            demoDtos.Add(new DemoDto() { Name = "Andrew", Desription = "測試文件二說明", Title = "測試文件二" });
            demoDtos.Add(new DemoDto() { Name = "Jacky", Desription = "測試文件三說明", Title = "測試文件三" });
            var docSource = Xceed.Words.NET.DocX.Load(@"D:\ITCT-F01-220128.docx");
            Xceed.Document.NET.Document docResult = null;
            foreach(var data in demoDtos)
            {
                var doc = docSource.Copy();
                doc.ReplaceText("{{ }}", data.Name);
                doc.ReplaceText("{{TITLE}}", data.Title);
                doc.ReplaceText("{{Desription}}", data.Desription);
                if (docResult == null)
                    docResult = doc;
                else
                    docResult.InsertDocument(doc);
            }
            docResult.SaveAs(@"D:\ITCT-F01-220128-1.docx");

        }
    }
    public class DemoDto
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public string Desription { get; set; }
    }
}
