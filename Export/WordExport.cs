using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;

namespace WordReportTest.Export
{
    public class WordExport
    {
        //Экспорт отчета в Word
        public void Export(List<ExportField> fields, string file, string template)
        {
            var doc = GenerateDocumentFromTemplate(file, template);

            var xDoc = doc.MainDocumentPart.GetXDocument();
            if ((xDoc.Root?.FirstNode) is XElement xBody)
            {
                CombineAttributeRun(xBody);

                doc.MainDocumentPart.Document.Body = new Body(xBody.ToString());
            }

            //fields.ForEach(f => f.Items.ForEach(fi =>
            //{
            //    if (fi.HasBackground)
            //    {
            //        SetBackground(doc, fi.AttrText, fi.Background);
            //        SetBackground(doc, fi.AttrDig, fi.Background);
            //    }

            //    SearchAndReplace(doc, fi.AttrText, fi.Content);
            //    SearchAndReplace(doc, fi.AttrDig, fi.Content);

            //    if(fi.IsUserField)
            //        SearchAndReplace(doc, fi.AttrUserField, fi.Content);
            //}));

            //ClearAttributes(doc);

            doc.Dispose();
        }

        private WordprocessingDocument GenerateDocumentFromTemplate(string file, string template)
        {
            File.Copy(template, file, true);

            var wordDoc = WordprocessingDocument.Open(file, true);
            wordDoc.ChangeDocumentType(WordprocessingDocumentType.Document); //Если на основе .dotx - меняем тип

            return wordDoc;
        }

        private void CombineAttributeRun(XElement xEl)
        {
            var ps = xEl.Descendants(W.p).ToList();
            foreach (var p in ps)
            {
                string contents = p.Descendants(W.t).Select(t => (string)t).StringConcatenate();
                if (contents.Contains("{") && contents.Contains("}"))
                {
                    int i1 = contents.IndexOf('{');
                    int i2 = contents.IndexOf('}');
                    if(i2 < i1) continue; //Error
                    var attr = contents.Substring(i1, i2 - i1 + 1);

                    var rs = p.Descendants(W.r).ToList();
                    //for


                    //rs.ForEach(r=>r.Remove());
                }
            }

        }

        private void SearchAndReplace(WordprocessingDocument doc, string search, string replace)
        {
            TextReplacer.SearchAndReplace(doc, search, replace, false);
        }

        private void SetBackground(WordprocessingDocument doc, string text, string back)
        {
            
        }

        private void ClearAttributes(WordprocessingDocument doc)
        {

        }


        public static string ParagraphText(XElement e)
        {
            XNamespace w = e.Name.Namespace;
            return e
                .Elements(w + "r")
                .Elements(w + "t")
                .StringConcatenate(element => (string)element);
        }

    }
}
