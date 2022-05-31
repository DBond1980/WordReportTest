using System.Collections.Generic;
using VT_Test.DataModels;

namespace WordReportTest.DataModels.ReportModel
{
    public class CycleResult
    {
        public int Id { get; set; }
        public string Terminals { get; set; }
        public string RatedSecondaryVoltage { get; set; }
        public string Class { get; set; }
        public string RatedPowerFactor { get; set; }
        public string RatedBurden1 { get; set; }
        public string RatedBurden2 { get; set; }
        public string RatedBurden3 { get; set; }
        public string Burden1 { get; set; }
        public string Burden2 { get; set; }
        public string Burden3 { get; set; }
        public string Thd { get; set; }
        public string ThdVolt { get; set; }
        public string ThdBack { get; set; }
        public string Asymmetry { get; set; }
        public string AsymmetryVolt { get; set; }
        public string AsymmetryBack { get; set; }

        //...................
        public int ReportId { get; set; }
        public Report Report { get; set; }
        public List<Result> Results { get; set; }


        public string GetTerminalBurden()
        {
            if (Report == null) return Burden1;
            if ((PrincipleEnum)Report.Principle == PrincipleEnum.InductiveThree)
            {
                if (Terminals == "c-a") return Burden3;
                if (Terminals == "b-c") return Burden2;
            }
            return Burden1;
        }

        public string GetTerminalRatedBurden()
        {
            if (Report == null) return RatedBurden1;
            if ((PrincipleEnum)Report.Principle == PrincipleEnum.InductiveThree)
            {
                if (Terminals == "c-a") return RatedBurden3;
                if (Terminals == "b-c") return RatedBurden2;
            }
            return RatedBurden1;
        }

    }


}
