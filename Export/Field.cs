using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using WordReportTest.DataModels.ReportModel;

namespace WordReportTest.Export
{
    public enum FieldType
    {//Типы полей
        Report,         //DataModel.ReportModel.Report
        CyclesGroup,    //DataModel.ReportModel.CycleResult...
        Result,         //DataModel.ReportModel.Result
        Samples,        //DataModel.ReportModel.Result.Samples...
    }
    public class Field
    {
        public static readonly List<Field> Fields = new List<Field>
        {
            new Field("ReportNumber",         1, FieldType.Report),
            new Field("Date",                 2, FieldType.Report),
            new Field("Time",                 3, FieldType.Report),
            new Field("DateTime",             4, FieldType.Report),
            new Field("Type",                 5, FieldType.Report),
            new Field("CardCode",             6, FieldType.Report),
            new Field("CardComment",          7, FieldType.Report),
            new Field("Manufacturer",         8, FieldType.Report),
            new Field("Year",                 9, FieldType.Report),
            new Field("Standard",            10, FieldType.Report),
            new Field("SerialNumber",        11, FieldType.Report),
            new Field("Owner",               12, FieldType.Report),
            new Field("Customer",            13, FieldType.Report),
            new Field("Substation",          14, FieldType.Report),
            new Field("RatedPrimaryVoltage", 15, FieldType.Report, true),
            new Field("RatedFrequency",      16, FieldType.Report, true),
            new Field("RatedVoltageFactor",  17, FieldType.Report, true),
            new Field("Humidity",            18, FieldType.Report, true),
            new Field("Temperature",         19, FieldType.Report, true),
            new Field("Conclusion",          20, FieldType.Report),
            new Field("TestedBy",            21, FieldType.Report),
            new Field("StateVerOfficer",     22, FieldType.Report),
            new Field("WarningZone",         23, FieldType.Report),
            new Field("UserFieldText1",      24, FieldType.Report),
            new Field("UserFieldText2",      25, FieldType.Report),
            new Field("UserFieldText3",      26, FieldType.Report),
            new Field("UserFieldText4",      27, FieldType.Report),
            new Field("UserFieldText5",      28, FieldType.Report),

            new Field("Terminals",            1, FieldType.CyclesGroup),
            new Field("Class",                2, FieldType.CyclesGroup, true),
            new Field("RatedBurden",          3, FieldType.CyclesGroup, true),
            new Field("RatedBurden1",         4, FieldType.CyclesGroup, true),
            new Field("RatedBurden2",         5, FieldType.CyclesGroup, true),
            new Field("RatedBurden3",         6, FieldType.CyclesGroup, true),
            new Field("Thd",                  7, FieldType.CyclesGroup, true, true),
            new Field("ThdVolt",              8, FieldType.CyclesGroup, true),
            new Field("Asymmetry",            9, FieldType.CyclesGroup, true, true),
            new Field("AsymmetryVolt",       10, FieldType.CyclesGroup, true),

            new Field("Voltage",              1, FieldType.Result, true),
            new Field("MeasVoltage",          2, FieldType.Result, true),
            new Field("Burden",               3, FieldType.Result, true),
            new Field("Burden1",              4, FieldType.Result, true),
            new Field("Burden2",              5, FieldType.Result, true),
            new Field("Burden3",              6, FieldType.Result, true),
            new Field("RatioError",           7, FieldType.Result, true, true),
            new Field("PhaseDisp",            8, FieldType.Result, true, true),
            new Field("Thd",                  9, FieldType.Result, true, true),
            new Field("Asymmetry",           10, FieldType.Result, true, true),
            new Field("Frequency",           11, FieldType.Result, true, true),

            new Field("RatioErrorSample",    12, FieldType.Samples, true),
            new Field("PhaseDispSample",     13, FieldType.Samples, true),
        };

        public Field(string name, int number, FieldType type, bool decSep = false, bool back = false)
        {
            Name = name;
            _number = number;
            _type = type;
            CanContainsDecSep = decSep;
            HasBackground = back;
        }


        public string Name { get; set; }
        public string NameDig => (((int)_type) + 1) + "." + _number;
        public bool CanContainsDecSep { get; set; }
        public bool HasBackground { get; set; }

        private int _number;
        private FieldType _type;

        public List<FieldItem> Items => _items;
        private List<FieldItem> _items = new List<FieldItem>();

        //Инициализация всех полей данными из отчета
        public static void InitFields(Report report)
        {

        }
    }
}
