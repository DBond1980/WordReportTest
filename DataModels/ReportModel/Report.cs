using System;
using System.Collections.Generic;

namespace WordReportTest.DataModels.ReportModel
{
    public class Report
    {
        public int Id { get; set; }

        public DateTime DateTime { get; set; }

        //............................................................................
        //Информация с тестовой карты
        //............................................................................
        public string Manufacturer { get; set; }
        public string Type { get; set; }
        public int Standard { get; set; }
        public string Comment { get; set; }
        public string RatedPrimaryVoltage { get; set; }
        public string RatedFrequency { get; set; }
        public string RatedVoltageFactor { get; set; }
        public int Principle { get; set; }
        public int Insulation { get; set; }
        public int MeasWind { get; set; }
        public bool ResidWind { get; set; }
        public bool TapWind { get; set; }
        public int TestingProgram { get; set; }
        public string WarningZone { get; set; }
        //Заложено на будущее развитие программы
        public string CardCode { get; set; } //Код карты (возможно как barcode)
        public string Limit { get; set; } //Ограничение максимального напряжения
                                          //(для предварительных испытаний на заводе)

        //............................................................................
        //Информация о конкретном трансформаторе и конкретном измерении
        //............................................................................
        public string Serial { get; set; }
        public string ReportNumber { get; set; }
        public string TestedBy { get; set; }
        public string StateVerificationOfficer { get; set; }
        public string Customer { get; set; }
        public string Owner { get; set; }
        public string YearOfManufacture { get; set; }
        public string Substation { get; set; }
        public string Temperature { get; set; }
        public string Humidity { get; set; }
        public string Conclusion { get; set; }

        //Поля, настраиваемые пользователем
        public string UserField1Name { get; set; }
        public string UserField1Content { get; set; }
        public string UserField2Name { get; set; }
        public string UserField2Content { get; set; }
        public string UserField3Name { get; set; }
        public string UserField3Content { get; set; }
        public string UserField4Name { get; set; }
        public string UserField4Content { get; set; }
        public string UserField5Name { get; set; }
        public string UserField5Content { get; set; }
        //............................................................................

        public List<CycleResult> CycleResults { get; set; }
    }

}
