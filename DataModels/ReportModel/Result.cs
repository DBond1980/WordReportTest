
namespace WordReportTest.DataModels.ReportModel
{
    public class Result
    {
        public int Id { get; set; }
        public string Voltage { get; set; }
        public string MeasVoltage { get; set; }
        public string RatioError { get; set; }
        public string RatioErrorBack { get; set; }
        public string PhaseDisp { get; set; }
        public string PhaseDispBack { get; set; }
        public string Thd { get; set; }
        public string Asymmetry { get; set; }
        public string Frequency { get; set; }
        public string RatioErrorSamples { get; set; }
        public string PhaseDispSamples { get; set; }

        //...................
        public int CycleResultId { get; set; }
        public CycleResult CycleResult { get; set; }
    }

}
