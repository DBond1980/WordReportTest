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

            //wordDoc.MainDocumentPart.AddImagePart(I)

            if ((xDoc.Root?.FirstNode) is XElement xBody)
            {
                CombineAttributeRun(xBody);

                WordExportTable.InitTables(xBody);
                WordExportTable.AutomaticRowNumbering();
                WordExportTable.AutomaticNumbering();

                if(RepeatingBlockCopy(xBody, ExportField.MeasGroupNum))
                    WordExportTable.InitTables(xBody); //Переинициализация таблиц т.к. могло изменится их количество

                WordExportTable.RemoveTableAttributes(xBody);

                fields.ForEach(f => f.Items.ForEach(fi =>
                {
                    if (fi.HasBackground)
                    {
                        SetBackground(xBody, fi.AttrText, fi.Background, false);
                        SetBackground(xBody, fi.AttrDig, fi.Background, false);
                    }
                }));

                fields.ForEach(f => f.Items.ForEach(fi =>
                {
                    SearchAndReplace(xBody, fi.AttrText, fi.Content);
                    SearchAndReplace(xBody, fi.AttrDig, fi.Content);

                    if (fi.IsUserField)
                        SearchAndReplace(xBody, fi.AttrUserField, fi.Content);
                }));

                WordExportTable.DeleteEmptyRows();
                WordExportTable.MergeTables();

                var xBodyStr = ClearAttributes(xBody.ToString());
                //var xBodyStr = xBody.ToString();

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
        //Каждый run может содержать только один атрибут
        private void CombineAttributeRun(XElement xEl)
        {
            var w = xEl.Name.Namespace;
            var xParagraphs = xEl.Descendants(w + "p").ToList();
            foreach (var xParagraph in xParagraphs)
            {
                var contents = xParagraph.Descendants(w + "t").Select(t => (string)t).StringConcatenate();
                if (!contents.Contains('{') || !contents.Contains('}')) continue;

                //Разбитые run, которые содержат фигурные скобки на runы содержащие один символ
                //для корректного дальнейшего группирования
                var xRuns = xParagraph.Descendants(w + "r").ToList();
                foreach (var xRun in xRuns)
                {
                    var xText = xRun.Element(w + "t");
                    if (xText == null) continue;
                    if ((xText.Value.Contains('{') || xText.Value.Contains('}'))
                        && xText.Value.Length > 1)
                    {
                        var text = xText.Value.ToCharArray();
                        foreach (var t in text)
                        {
                            var xRunNew = new XElement(xRun);
                            var xTextNew = xRunNew.Element(w + "t");
                            if (xTextNew != null)
                            {
                                if (xTextNew.Value[0] == ' ' || xTextNew.Value[xTextNew.Value.Length - 1] == ' ')
                                    xTextNew.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                                else
                                    xTextNew.Attribute(XNamespace.Xml + "space")?.Remove();

                                xTextNew.Value = t.ToString();
                            }
                            xRun.AddBeforeSelf(xRunNew);
                        }
                        xRun.Remove();
                    }
                }

                //Группирование run
                xRuns = xParagraph.Descendants(w + "r").ToList();
                XElement xTextAttr = null;
                foreach (var xRun in xRuns)
                {
                    var xText = xRun.Element(w + "t");
                    if (xText == null) continue;
                    if (xText.Value.Contains('{'))
                    {
                        xText.Value = xText.Value.ToUpper();
                        if (xText.Value.Contains('}')) { xTextAttr = null; continue; }
                        xTextAttr = xText;
                        xTextAttr.Attribute(XNamespace.Xml + "space")?.Remove();
                        continue;
                    }

                    if (xTextAttr != null)
                    {
                        xTextAttr.Value += xText.Value.ToUpper();
                        xRun.Remove();
                    }

                    if (xText.Value.Contains('}')) xTextAttr = null;
                }
            }
        }

        //Копирование повторяющегося блока с заменой в атрибутах N на номер блока
        private bool RepeatingBlockCopy(XElement xEl, int copyNumber)
        {
            var w = xEl.Name.Namespace;

            XElement beginElement = null;
            XElement endElement = null;
            int beginIndex = -1;
            int endIndex = -1;

            //Поиск начала и конца повторяющегося блока
            var xElements = xEl.Elements().ToList();
            foreach (var xElement in xElements)
            {
                if(xElement.Name != w + "p") continue;

                var contents = xElement.Descendants(w + "t").Select(t => (string)t).StringConcatenate();
                if (contents.ToUpper().Contains("{RepeatingBlock.Begin}".ToUpper()))
                {
                    beginElement = xElement;
                    beginIndex = xElements.IndexOf(beginElement);
                }
                if (contents.ToUpper().Contains("{RepeatingBlock.End}".ToUpper()))
                {
                    endElement = xElement;
                    endIndex = xElements.IndexOf(endElement);
                }
            }

            if (beginElement != null && endElement != null && beginIndex < endIndex)
            {
                var repeatElements = xElements.GetRange(beginIndex + 1, endIndex - beginIndex - 1).ToList();

                for (int i = 0; i < copyNumber; i++)
                {
                    foreach (var repeatElement in repeatElements)
                    {
                        var repeatElementCopy = new XElement(repeatElement);
                        SearchAndReplace(repeatElementCopy, "{N/", "{" + (i + 1) + "/");
                        SearchAndReplace(repeatElementCopy, "{N.", "{" + (i + 1) + ".");
                        endElement.AddBeforeSelf(repeatElementCopy);
                    }
                }

                beginElement.Remove();
                endElement.Remove();
                repeatElements.Remove();

                return true;
            }

            return false;
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

        //Поиск атрибутов и их замена
        private void SearchAndReplace(XElement xEl, string search, string replace)
        {
            var w = xEl.Name.Namespace;

            search = search.ToUpper();

            var xParagraphs = xEl.Name != w + "p" ? 
                xEl.Descendants(w + "p").ToList() : 
                new List<XElement>(new []{ xEl });

            foreach (var xParagraph in xParagraphs)
            {
                var contents = xParagraph.Descendants(w + "t").Select(t => (string)t).StringConcatenate();
                if (!contents.Contains(search)) continue;

                var xRuns = xParagraph.Descendants(w + "r").ToList();

                foreach (var xRun in xRuns)
                {
                    var xText = xRun.Element(w + "t");
                    if (xText == null) continue;
                    if (xText.Value.Contains(search))
                    {
                        xText.Value = xText.Value.Replace(search.ToUpper(), replace);
                    }
                }
            }
            //return str.Replace(search.ToUpper(), replace);
            //TextReplacer.SearchAndReplace(doc, search, replace, false);
        }

        //Выделение элементов выходящих за допустимые границы или приближаются к ним
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
