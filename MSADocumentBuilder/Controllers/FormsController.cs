using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Hosting;
using System.Web;



using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MSADocumentBuilder.Models;
using System.Text.RegularExpressions;

namespace MSADocumentBuilder.Controllers
{
    public class FormsController : Controller
    {
        private IHostingEnvironment _env;
        public FormsController(IHostingEnvironment env)
        {
            _env = env;
        }
        public IActionResult Index()

        {
            return View();
        }

        [HttpPost]
        public IActionResult MSA(MSAModel msa)
        {

            //string strDoc = @"C:\Users\jakmoz01\Documents\MSADocumentBuilder\SampleMSA.docx";
            var dir = _env.WebRootPath;
            String strDoc = dir + "\\Documents\\SampleMSA.docx";

            SearchAndReplace(strDoc,msa);
            return View();

        }

        private void OpenAndAddToWordprocessingStream(Stream stream, string txt)
        {
            // Open a WordProcessingDocument based on a stream.
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(stream, true);

            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(txt));

            // Close the document handle.
            wordprocessingDocument.Close();

            // Caller must close the stream.
        }
        static void SearchAndReplace(string document, MSAModel msa)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex("Insert Date");
                docText = regexText.Replace(docText, msa.msaDate);

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

    }
}