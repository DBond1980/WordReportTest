using System;
using System.Collections.Generic;
using System.Linq;
using WordReportTest.DataModels.ReportModel;

namespace WordReportTest.Export
{
    public class ExportField
    {
        public static readonly List<ExportField> ExportFields = new List<ExportField>
        {
            //Report Fields -> {ReportNumber}, {1}
            new ExportField("ReportNumber",         1),
            new ExportField("Date",                 2),
            new ExportField("Time",                 3),
            new ExportField("DateTime",             4),
            new ExportField("Type",                 5),
            new ExportField("CardComment",          6),
            new ExportField("Manufacturer",         7),
            new ExportField("Year",                 8),
            new ExportField("Standard",             9),
            new ExportField("SerialNumber",        10),
            new ExportField("Owner",               11),
            new ExportField("Customer",            12),
            new ExportField("Substation",          13),
            new ExportField("RatedPrimaryVoltage", 14, true),
            new ExportField("RatedFrequency",      15, true),
            new ExportField("RatedVoltageFactor",  16, true),
            new ExportField("Humidity",            17, true),
            new ExportField("Temperature",         18, true),
            new ExportField("Conclusion",          19),
            new ExportField("TestedBy",            20),
            new ExportField("StateVerOfficer",     21),
            new ExportField("WarningZone",         22),
            new ExportField("UserField1",          23),
            new ExportField("UserField2",          24),
            new ExportField("UserField3",          25),
            new ExportField("UserField4",          26),
            new ExportField("UserField5",          27),

            //CyclesGroup Fields -> {1/Terminals}, {1/1}
            new ExportField("Terminals",            1),
            new ExportField("Class",                2, true),
            new ExportField("RatedBurden",          3, true),
            new ExportField("RatedBurden1",         4, true),
            new ExportField("RatedBurden2",         5, true),
            new ExportField("RatedBurden3",         6, true),
            new ExportField("Thd",                  7, true, true),
            new ExportField("ThdVolt",              8, true),
            new ExportField("Asymmetry",            9, true, true),
            new ExportField("AsymmetryVolt",       10, true),

            //Result Fields -> {1.1/Voltage}, {1.1/1}
            new ExportField("Voltage",              1, true),
            new ExportField("MeasVoltage",          2, true),
            new ExportField("Burden",               3, true),
            new ExportField("Burden1",              4, true),
            new ExportField("Burden2",              5, true),
            new ExportField("Burden3",              6, true),
            new ExportField("RatioError",           7, true, true),
            new ExportField("PhaseDisp",            8, true, true),
            new ExportField("Thd",                  9, true),
            new ExportField("Asymmetry",           10, true),
            new ExportField("Frequency",           11, true),

            //Samples Fields -> {1.1.1/RatioErrorSample}, {1.1.1/1}
            new ExportField("RatioErrorSample",    1, true),
            new ExportField("PhaseDispSample",     2, true)
        };

        public static int MeasGroupNum = 0;

        public ExportField(string name, int number, bool decSep = false, bool back = false)
        {
            Name = name;
            _number = number;
            //_type = type;
            CanContainsDecSep = decSep;
            HasBackground = back;
        }


        public string Name { get; set; }
        //public string NameDig => (((int)_type) + 1) + "." + _number;
        public string NameDig => _number.ToString();
        public bool CanContainsDecSep { get; set; }
        public bool HasBackground { get; set; }

        private int _number;
        //private FieldType _type;

        public List<ExportFieldItem> Items => _items;
        private List<ExportFieldItem> _items = new List<ExportFieldItem>();

        //Инициализация всех полей данными из отчета
        public static void InitFields(Report r)
        {
            ClearFieldItems();

            SetField("ReportNumber", r.ReportNumber);
            SetField("Date", r.DateTime.ToString("dd.MM.yyyy"));
            SetField("Time", r.DateTime.ToString("HH:mm:ss"));
            SetField("DateTime", r.DateTime.ToString("dd.MM.yyyy HH:mm:ss"));
            SetField("Type", r.Type);
            SetField("CardComment", r.Comment);
            SetField("Manufacturer", r.Manufacturer);
            SetField("Year", r.YearOfManufacture);
            SetField("Standard", r.Standard.ToString()); //!!!!!!! доработать
            SetField("SerialNumber", r.Serial);
            SetField("Owner", r.Owner);
            SetField("Customer", r.Customer);
            SetField("Substation", r.Substation);
            SetField("RatedPrimaryVoltage", r.RatedPrimaryVoltage);
            SetField("RatedFrequency", r.RatedFrequency);
            SetField("RatedVoltageFactor", r.RatedVoltageFactor);
            SetField("Humidity", r.Humidity);
            SetField("Temperature", r.Temperature);
            SetField("Conclusion", r.Conclusion);
            SetField("TestedBy", r.TestedBy);
            SetField("StateVerOfficer", r.StateVerificationOfficer);
            SetField("WarningZone", r.WarningZone);
            SetField("UserField1", r.UserField1Content, r.UserField1Name);
            SetField("UserField2", r.UserField2Content, r.UserField2Name);
            SetField("UserField3", r.UserField3Content, r.UserField3Name);
            SetField("UserField4", r.UserField4Content, r.UserField4Name);
            SetField("UserField5", r.UserField5Content, r.UserField5Name);

            //Объединение циклов по Terminals и Class
            int id1 = 0;
            var cycleGroupList = r.CycleResults.GroupBy(c => c.Terminals + c.Class, t => t).ToList();
            MeasGroupNum = cycleGroupList.Count;
            foreach (var cycleGroup in cycleGroupList)
            {
                id1++;
                int id2 = 0;
                foreach (var c in cycleGroup)
                {
                    SetField("Terminals", c.Terminals, id1);
                    SetField("Class", c.Class, id1);
                    SetField("RatedBurden", c.GetTerminalRatedBurden(), id1);
                    SetField("RatedBurden1", c.RatedBurden1, id1);
                    SetField("RatedBurden2", c.RatedBurden2, id1);
                    SetField("RatedBurden3", c.RatedBurden3, id1);
                    SetField("Thd", c.Thd, id1, back: c.ThdBack);
                    SetField("ThdVolt", c.Thd, id1);
                    SetField("Asymmetry", c.Thd, id1, back: c.AsymmetryBack);
                    SetField("AsymmetryVolt", c.Thd, id1);

                    foreach (var result in c.Results)
                    {
                        id2++;
                        int id3 = 0;

                        SetField("Voltage", result.Voltage, id1, id2);
                        SetField("MeasVoltage", result.MeasVoltage, id1, id2);
                        SetField("Burden", c.GetTerminalBurden(), id1, id2);
                        SetField("Burden1", c.Burden1, id1, id2);
                        SetField("Burden2", c.Burden2, id1, id2);
                        SetField("Burden3", c.Burden3, id1, id2);
                        SetField("RatioError", result.RatioError, id1, id2, back: result.RatioErrorBack);
                        SetField("PhaseDisp", result.PhaseDisp, id1, id2, back: result.PhaseDispBack);
                        SetField("Thd", result.Thd, id1, id2);
                        SetField("Asymmetry", result.Asymmetry, id1, id2);
                        SetField("Frequency", result.Frequency, id1, id2);
                        
                        result.RatioErrorSamples.Split(';').ToList().ForEach(s =>
                        {
                            if(!string.IsNullOrWhiteSpace(s)) 
                                SetField("RatioErrorSample", s, id1, id2, ++id3);
                        });

                        id3 = 0;

                        result.PhaseDispSamples.Split(';').ToList().ForEach(s =>
                        {
                            if (!string.IsNullOrWhiteSpace(s))
                                SetField("PhaseDispSample", s, id1, id2, ++id3);
                        });
                    }
                }
            }
        }

        private static void SetField(string name, string content, int id1, int id2 = -1, int id3 = -1, string back = null)
        {
            var field = ExportFields.FirstOrDefault(f => f.Name == name);
            if (field == null) throw new Exception();
            field.Items.Add(new ExportFieldItem(field, content, id1, id2, id3, back: back));
        }

        private static void SetField(string name, string content, string userFieldName = null)
        {
            var field = ExportFields.FirstOrDefault(f => f.Name == name);
            if(field == null) throw new Exception();
            field.Items.Add(new ExportFieldItem(field, content, userFieldName: userFieldName));
        }

        private static void ClearFieldItems()
        {
            ExportFields.ForEach(f => f.Items.Clear());
        }
    }
}
