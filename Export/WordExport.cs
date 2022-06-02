using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
            var wordDoc = GenerateDocumentFromTemplate(file, template);

            var xDoc = wordDoc.MainDocumentPart.GetXDocument();

            if ((xDoc.Root?.FirstNode) is XElement xBody)
            {
                CombineAttributeRun(xBody);


                fields.ForEach(f => f.Items.ForEach(fi =>
                {
                    if (fi.HasBackground)
                    {
                        SetBackground(xBody, fi.AttrText, fi.Background, false);
                        SetBackground(xBody, fi.AttrDig, fi.Background, false);
                    }
                }));

                var xBodyStr = xBody.ToString();

                fields.ForEach(f => f.Items.ForEach(fi =>
                {
                    xBodyStr = SearchAndReplace(xBodyStr, fi.AttrText, fi.Content);
                    xBodyStr = SearchAndReplace(xBodyStr, fi.AttrDig, fi.Content);

                    if (fi.IsUserField)
                        xBodyStr = SearchAndReplace(xBodyStr, fi.AttrUserField, fi.Content);

                }));

                xBodyStr = ClearAttributes(xBodyStr);

                wordDoc.MainDocumentPart.Document.Body = new Body(xBodyStr);
            }

            wordDoc.Dispose();
        }

        //Создание документа на основе шаблона
        private WordprocessingDocument GenerateDocumentFromTemplate(string file, string template)
        {
            File.Copy(template, file, true);

            var wordDoc = WordprocessingDocument.Open(file, true);
            wordDoc.ChangeDocumentType(WordprocessingDocumentType.Document); //Если на основе .dotx - меняем тип

            return wordDoc;
        }

        //Подготовка атрибутов, т.е. перенос каждого атрибута в один run и преобразование в заглавные буквы,
        //для легкой и быстрой дальнейшей замены (после этого можно заменять преобразовав в файл в строку)
        private void CombineAttributeRun(XElement xEl)
        {
            var w = xEl.Name.Namespace;
            var xParagraphs = xEl.Descendants(w + "p").ToList();
            foreach (var xParagraph in xParagraphs)
            {
                var contents = xParagraph.Descendants(w + "t").Select(t => (string)t).StringConcatenate();
                if (!contents.Contains("{") || !contents.Contains("}")) continue;

                var xRuns = xParagraph.Descendants(w + "r").ToList();
                XElement xTextAttr = null;
                foreach (var xRun in xRuns)
                {
                    var xText = xRun.Element(w + "t");
                    if (xText == null) continue;
                    if (xText.Value.Contains('{'))
                    {
                        xText.Value = xText.Value.ToUpper();
                        if (xText.Value.Contains('}')) { xTextAttr = null; continue;}
                        xTextAttr = xText;
                        continue;
                    }

                    if (xTextAttr!=null)
                    {
                        xTextAttr.Value += xText.Value.ToUpper();
                        xRun.Remove();
                    }

                    if (xText.Value.Contains('}')) xTextAttr = null;
                }
            }
        }

        //Поиск и удаление оставшихся атрибутов
        private string ClearAttributes(string str)
        {
            // @"{([^\{\}]*)}" - все символы заключенные в фигурные скобки 
            //                   если туда не входят сами фигурные скобки
            str = (new Regex(@"{([^\{\}]*)}")).Replace(str, "");
            return str;
        }

        //Поиск атрибутов и их замена
        private string SearchAndReplace(string str, string search, string replace)
        {
            return str.Replace(search.ToUpper(), replace);
            //TextReplacer.SearchAndReplace(doc, search, replace, false);
        }

        private void SetBackground(XElement xEl, string text, string back, bool isBlackWhite)
        {
            if(string.IsNullOrWhiteSpace(back)) return;
            back = back.ToUpper();
            if (back == "GREEN") return;

            var w = xEl.Name.Namespace;
            var xParagraphs = xEl.Descendants(w + "p").ToList();

            foreach (var xParagraph in xParagraphs)
            {
                var contents = xParagraph.Descendants(w + "t").Select(t => (string)t).StringConcatenate();
                if (!contents.Contains(text)) continue;

                var xRuns = xParagraph.Descendants(w + "r").ToList();
                foreach (var xRun in xRuns)
                {
                    var xText = xRun.Element(w + "t");
                    if(xText == null) continue;
                    if (xText.Value.Contains(text))
                    {
                        //Подчеркивание run
                        var xRunProp = xRun.Element(w + "rPr");
                        AddOrChangeXEl(xRunProp, "u", "val",
                            back == "RED" ? "single" : "dotted");

                        if (!isBlackWhite)
                        {
                            if (xParagraph.Parent?.Name == w + "tc")
                            {//Закрасить ячейку если параграф находится в таблице
                                var xTProp = xParagraph.Parent?.Element(w + "tcPr");
                                AddOrChangeXEl(xTProp, "shd","fill",
                                    back == "RED" ? "FF7D7D" : "FFC000");
                            }
                            else
                            {//Закрасить текст если не находится в таблице
                                AddOrChangeXEl(xRunProp, "shd", "fill",
                                    back == "RED" ? "FF7D7D" : "FFC000");
                            }
                        }
                    }
                }
            }
        }

        //Добавление или изменение элемента с атрибутом
        private void AddOrChangeXEl(XElement xElParent, string xElName, string xAttrName, string xAttrValue)
        {
            if(xElParent == null) return;
            var w = xElParent.Name.Namespace;
            var xEl = xElParent.Element(w + xElName);
            if (xEl == null)
            {//Не существует -> создаем новый
                xElParent.Add(new XElement(w + xElName, new XAttribute(w + xAttrName, xAttrValue)));
            }
            else
            {//Элемент существует -> добавляет атрибут
                xEl.SetAttributeValue(w + xAttrName, xAttrValue);
            }
        }
    }


}
