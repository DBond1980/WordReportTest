using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using OpenXmlPowerTools;

namespace WordReportTest.Export
{
    public class WordExportTable
    {
        private XElement _xTable;
        private bool _propDeleteEmptyRows = false;
        private bool _propAutoNum = false;
        private int _propAutoNumN = 0;
        private bool _propAutoRowNum = false;
        private bool _propContinue = false;
        private static XNamespace w;

        private static List<WordExportTable> _tables = new List<WordExportTable>();

        public static void InitTables(XElement xEl)
        {
            _tables.Clear();

            w = xEl.Name.Namespace;
            var xParagraphs = xEl.Descendants(w + "p").ToList();

            foreach (var xParagraph in xParagraphs)
            {
                var contents = xParagraph.Descendants(w + "t").Select(t => (string) t).StringConcatenate();
                if (!contents.Contains("{#TABLE")) continue;
                if (xParagraph.NextNode is XElement tbl)
                {
                    if (tbl.Name == w + "tbl")
                    {
                        _tables.Add(new WordExportTable(contents, tbl));
                    }
                    xParagraph.Remove();
                }
            }
        }

        public WordExportTable(string propStr, XElement xTbl)
        {
            _xTable = xTbl;

            //Извлекаем атрибут
            propStr = (new Regex(@"{#TABLE([^\{\}]*)}")).Match(propStr).Value;
            if (string.IsNullOrWhiteSpace(propStr)) return;

            //Инициализируем свойства
            var props = propStr.Split('/', '\\', '(', ')', '[', ']', '{', '}');

            //Удаление пустых рядов из таблицы
            _propDeleteEmptyRows = props.FirstOrDefault(p => p.Contains("DELETEEMPTYROWS")) != null;

            //Автоматическое добавление номера таблицы в отчете
            //т.е. если AutoNum(2) то {1/7} -> {2.1/7} и {3.5/2} -> {2.2.5/2)
            //или если AutoNum(n) то {1/7} -> {n.1/7} и {3.5/2} -> {n.2.5/2)
            var autoNumIndex = props.ToList().FindIndex(p => p.Contains("AUTONUM"));
            if (autoNumIndex >= 0)
            {
                _propAutoNum = true;
                if (++autoNumIndex < props.Length)
                {
                    if (!int.TryParse(props[autoNumIndex], out _propAutoNumN))
                        _propAutoNumN = 0;
                }
                else _propAutoNumN = 0;
            }
            else _propAutoNum = false;

            //Автоматическое добавление номера ряда
            //т.е. ко всем атрибутам добавляется номер ряда (или, точнее, добавляется число повторений одинаковых атрибутов)
            _propAutoRowNum = props.FirstOrDefault(p => p == "AUTOROWNUM") != null;

            //Продолжение таблицы
            _propContinue = props.FirstOrDefault(p => p == "CONTINUE") != null;
        }

        //Соединение таблиц
        public static void MergeTables()
        {
            XElement mainTable = null;
            foreach (var wordExportTable in _tables)
            {
                if (!wordExportTable._propContinue)
                {
                    mainTable = wordExportTable._xTable;
                    continue;
                }

                var trs = wordExportTable._xTable.Descendants(w + "tr");
                mainTable?.Add(trs);
                wordExportTable._xTable.Remove();
            }
        }


        //Автоматическое добавление номера таблицы в отчете
        //т.е. если AutoNum(2) то {1/7} -> {2.1/7} и {3.5/2} -> {2.2.5/2)
        //или если AutoNum(n) то {1/7} -> {n.1/7} и {3.5/2} -> {n.2.5/2)
        public static void AutomaticNumbering()
        {
            foreach (var wordExportTable in _tables)
            {
                if (!wordExportTable._propAutoNum) continue;

                var xTexts = wordExportTable._xTable.Descendants(w + "t").ToList()
                    .FindAll(xt => xt.Value.Contains('{') || xt.Value.Contains('}'));

                //Поиск всех атрибутов, к которым нужно добавить номер
                var attrList = new List<string>();

                foreach (var xText in xTexts)
                {
                    var match = (new Regex(@"{([^\{\}]*)}")).Match(xText.Value);
                    if (!attrList.Contains(match.Value) && !match.Value.Contains('$'))
                        attrList.Add(match.Value);
                }

                //Добавления номера в атрибут
                var xTextChangeList = new List<XElement>();
                foreach (var attr in attrList)
                {
                    var nStr = wordExportTable._propAutoNumN == 0 ? "N" : wordExportTable._propAutoNumN.ToString();
                    foreach (var xText in xTexts)
                    {
                        if (xText.Value.Contains(attr) && !xTextChangeList.Contains(xText))
                        {
                            var attrReplace = attr.Insert(1, attr[1] == '/' ? nStr : nStr + ".");
                            xText.Value = xText.Value.Replace(attr, attrReplace);
                            xTextChangeList.Add(xText);
                        }
                    }
                }

                //Удаление '$'
                foreach (var xText in xTexts)
                    xText.Value = xText.Value.Replace("{$", "{");
            }
        }

        //Автоматическое добавление номера ряда
        //т.е. ко всем атрибутам добавляется номер ряда (или, другими словами, добавляется число повторений одинаковых атрибутов)
        //Если атрибут содержит символ '$' он не изменяется, а '$' удаляется
        public static void AutomaticRowNumbering()
        {
            foreach (var wordExportTable in _tables)
            {
                if (!wordExportTable._propAutoRowNum) continue;

                var xTexts = wordExportTable._xTable.Descendants(w + "t").ToList()
                    .FindAll(xt => xt.Value.Contains('{') || xt.Value.Contains('}'));

                //Поиск всех атрибутов, к которым нужно добавить номер
                var attrList = new List<string>();

                foreach (var xText in xTexts)
                {
                    var match = (new Regex(@"{([^\{\}]*)}")).Match(xText.Value);
                    if (!attrList.Contains(match.Value) && !match.Value.Contains('$'))
                        attrList.Add(match.Value);
                }

                //Добавления номера в атрибут
                var xTextChangeList = new List<XElement>();
                foreach (var attr in attrList)
                {
                    var n = 1;
                    foreach (var xText in xTexts)
                    {
                        if (xText.Value.Contains(attr) && !xTextChangeList.Contains(xText))
                        {
                            var attrReplace = attr.Insert(1, attr[1] == '/' ? n.ToString() : n + ".");
                            xText.Value = xText.Value.Replace(attr, attrReplace);
                            n++;
                            xTextChangeList.Add(xText);
                        }
                    }
                }

                //Удаление '$'
                foreach (var xText in xTexts)
                    xText.Value = xText.Value.Replace("{$", "{");
            }
        }

        //Удаление пустых рядов из таблицы
        public static void DeleteEmptyRows()
        {
            foreach (var wordExportTable in _tables)
            {
                if(!wordExportTable._propDeleteEmptyRows) continue;

                var xRows = wordExportTable._xTable.Descendants(w + "tr").ToList();

                foreach (var xRow in xRows)
                {
                    var xCells = xRow.Descendants(w + "tc").ToList();
                    var isEmpty = true;
                    foreach (var xCell in xCells)
                    {
                        var contents = xCell.Descendants(w + "t").Select(t => (string)t).StringConcatenate();

                        if ((contents.Contains('{') && contents.Contains('}')) ||
                            string.IsNullOrWhiteSpace(contents)) continue;

                        isEmpty = false;
                        break;
                    }
                    if (isEmpty) xRow.Remove(); 
                }
            }
        }
    }
}
