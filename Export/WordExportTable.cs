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
            //т.е. ко всем атрибутам добавляется номер ряда (или, другими словами, добавляется число повторений одинаковых атрибутов)
            _propAutoRowNum = props.FirstOrDefault(p => p == "AUTOROWNUM") != null;
        }

        //Автоматическое добавление номера ряда
        //т.е. ко всем атрибутам добавляется номер ряда (или, другими словами, добавляется число повторений одинаковых атрибутов)
        public static void AutomaticRowNumbering()
        {
            foreach (var wordExportTable in _tables)
            {
                if (!wordExportTable._propAutoRowNum) continue;

                var xTexts = wordExportTable._xTable.Descendants(w + "t").ToList()
                    .FindAll(xt => xt.Value.Contains('{') || xt.Value.Contains('}'));

                while (xTexts.Count > 0)
                {
                    var attr = (new Regex(@"{([^\{\}]*)}")).Match(xTexts[0].Value).Value;
                    //var newAttr = 
                    //var stIndex = xTexts[0].Value.IndexOf('{');

                }

                //foreach (var xRow in xRows)
                //{
                //    var xCells = xRow.Descendants(w + "tc").ToList();
                    //var isEmpty = true;
                    //foreach (var xCell in xCells)
                    //{
                    //    var contents = xCell.Descendants(w + "t").Select(t => (string)t).StringConcatenate();

                    //    if ((contents.Contains('{') && contents.Contains('}')) ||
                    //        string.IsNullOrWhiteSpace(contents)) continue;

                    //    isEmpty = false;
                    //    break;
                    //}
                    //if (isEmpty) xRow.Remove();
                //}
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
