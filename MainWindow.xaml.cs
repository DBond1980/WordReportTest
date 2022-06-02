using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using WordReportTest.DataModels.ReportModel;
using WordReportTest.Export;

namespace WordReportTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        private void ButtonStart_Click(object sender, RoutedEventArgs e)
        {
            var r = GetReport();
            ExportField.InitFields(r);


            var wordExport = new WordExport();
            wordExport.Export(ExportField.ExportFields,
                @"C:\Work\C#\WordReportTest\!DOCX\Report1.docx",
                @"C:\Work\C#\WordReportTest\!DOCX\TempDocx.docx");



            //File.Copy(@"C:\Work\C#\WordReportTest\!DOCX\TempDocx.docx",
            //    @"C:\Work\C#\WordReportTest\!DOCX\Report1.docx", true);

            //using (var wordDoc = WordprocessingDocument.Open(@"C:\Work\C#\WordReportTest\!DOCX\Report1.docx", true))
            //{
            //    wordDoc.ChangeDocumentType(WordprocessingDocumentType.Document); //Если на основе .dotx - меняем тип

            //    TextReplacer.SearchAndReplace(wordDoc, "{Manufacturer}", "OLTEST", false);
            //    RepeatingBlockCopy(wordDoc);

                
            //}

        }


        //static void GenerateDocumentFromTemplate(string inputPath, string outputPath)
        //{
        //    MemoryStream documentStream;
        //    using (Stream stream = File.OpenRead(inputPath))
        //    {
        //        documentStream = new MemoryStream((int)stream.Length);
        //        //CopyStream(stream, documentStream);
        //        stream.CopyTo(documentStream);
        //        documentStream.Position = 0L;
        //    }

        //    using (WordprocessingDocument template = WordprocessingDocument.Open(documentStream, true))
        //    {
        //        template.ChangeDocumentType(WordprocessingDocumentType.Document);
        //        MainDocumentPart mainPart = template.MainDocumentPart;
        //        mainPart.DocumentSettingsPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate",
        //            new Uri(inputPath, UriKind.Absolute));

        //        mainPart.Document.Save();
        //    }


        //    File.WriteAllBytes(outputPath, documentStream.ToArray());
        //}




        private void RepeatingBlockCopy(WordprocessingDocument wordDoc)
        {
            var xDoc = wordDoc.MainDocumentPart.GetXDocument();

            if ((xDoc.Root?.FirstNode) is XElement xBody)
            {
                XElement beginElement = null;
                XElement endElement = null;
                int beginIndex = -1;
                int endIndex = -1;


                var xBodyList = xBody.Elements().ToList();

                foreach (var element in xBodyList)
                {
                    if (element.Name == W.p)
                    {
                        string contents = element.Descendants(W.t).Select(t => (string)t).StringConcatenate();
                        if (contents.ToUpper().Contains("{RepeatingBlock.Begin}".ToUpper()))
                        {
                            beginElement = element;
                            beginIndex = xBodyList.IndexOf(beginElement);
                        }
                        if (contents.ToUpper().Contains("{RepeatingBlock.End}".ToUpper()))
                        {
                            endElement = element;
                            endIndex = xBodyList.IndexOf(endElement);
                        }
                    }
                }

                if (beginElement != null && endElement != null && beginIndex < endIndex)
                {
                    var repeatElements = xBodyList.GetRange(beginIndex + 1, endIndex - beginIndex - 1);

                    foreach (var repeatElement in repeatElements)
                    {
                        //endElement.AddBeforeSelf(new XElement(repeatElement));
                        endElement.AddBeforeSelf(repeatElement);
                    }
                    

                    //xBodyList.InsertRange(endIndex-1, repeatElements.Select(e => new XElement(e)));
                }

                wordDoc.MainDocumentPart.Document.Body = new Body(xBody.ToString());

                //xDoc.Elements().First().ReplaceWith(xBody);

                //xDoc.Root?.
            }



            //if(xDoc.Root)
            //string contents = element.Descendants(W.t).Select(t => (string)t).StringConcatenate();


        }

        public static void SearchAndReplace(string document)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex("Hello world!");
                docText = regexText.Replace(docText, "Hi Everyone!");

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        public static void SearchAndReplace2(string document)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                TextReplacer.SearchAndReplace(wordDoc, "Hello world!", "Hi Everyone!", false);
            }
        }

        private Report GetReport()
        {
            var rnd = new Random();

            var r = 100;

            var report = new Report
            {
                DateTime = DateTime.Now,
                Manufacturer = "Производитель" + r,
                Type = "Тип1234567890" + r,
                Standard = 1000 + r,
                Comment = "Comment1234567890" + r,
                RatedPrimaryVoltage = "10/V3_r" + r,
                RatedFrequency = "50",
                RatedVoltageFactor = "1.2",
                Principle = 0,
                Insulation = 1,
                MeasWind = 2,
                ResidWind = true,
                TapWind = false,
                TestingProgram = 0,
                WarningZone = "80",
                CardCode = r.ToString(),
                Limit = "100",
                Serial = r.ToString(),
                ReportNumber = r.ToString(),
                TestedBy = "Бондарчук Д.А",
                StateVerificationOfficer = "Носко С.",
                Customer = "Олтест заказ",
                Owner = "Олтест собст.",
                YearOfManufacture = "1980",
                Substation = "П/С Троещина",
                Temperature = "20",
                Humidity = "50",
                Conclusion = "Все прекрасно " + r,
                UserField1Name = "UserField1Name" + r,
                UserField1Content = "UserField1Content" + r,
                UserField2Name = "UserField2Name" + r,
                UserField2Content = "UserField2Content" + r,
                UserField3Name = "UserField3Name" + r,
                UserField3Content = "UserField3Content" + r,
                UserField4Name = "UserField4Name" + r,
                UserField4Content = "UserField4Content" + r,
                UserField5Name = "UserField4Name" + r,
                UserField5Content = "UserField4Content" + r,
                CycleResults = new List<CycleResult>()
            };
            //reports.Add(report);

            var term = new[] { "a-b", "a-b", "b-c", "b-c", "c-a", "c-a" };
            var burden = new[] { "10", "2,5", "10", "2,5", "10", "2,5" };
            var voltage = new[] { "80", "100", "120" };
            var mVoltage = new[] { "80,1", "100,2", "120,3" };

            for (int rt = 1; rt <= 6; rt++)
            {

                var resultTable = new CycleResult
                {
                    Terminals = term[rt-1],
                    RatedSecondaryVoltage = "100 " + r + "_" + rt,
                    Class = "0,5",
                    RatedPowerFactor = "0.8",
                    RatedBurden1 = burden[0],
                    RatedBurden2 = burden[0],
                    RatedBurden3 = burden[0],
                    Burden1 = burden[rt - 1],
                    Burden2 = burden[rt - 1],
                    Burden3 = burden[rt - 1],
                    Thd = (2 + 10 * (rnd.NextDouble() * 2 - 1)).ToString("G3"),
                    ThdVolt = (100 + 10 * (rnd.NextDouble() * 2 - 1)).ToString("F0"),
                    ThdBack = "Green",
                    Asymmetry = (2 + 10 * (rnd.NextDouble() * 2 - 1)).ToString("G3"),
                    AsymmetryVolt = (120 + 10 * (rnd.NextDouble() * 2 - 1)).ToString("F0"),
                    AsymmetryBack = "Orange",
                    Results = new List<Result>(),
                    Report = report
                };
                report.CycleResults.Add(resultTable);

                for (int rr = 1; rr <= 3; rr++)
                {
                    var resultRow = new Result
                    {
                        Voltage = voltage[rr-1],
                        MeasVoltage = mVoltage[rr-1],
                        RatioError = (0.2 + 0.5 * (rnd.NextDouble() * 2 - 1)).ToString("G3"),
                        RatioErrorBack = "Orange",
                        PhaseDisp = (2 + 10 * (rnd.NextDouble() * 2 - 1)).ToString("G3"),
                        PhaseDispBack = "Red",
                        Thd = (2 + 10 * (rnd.NextDouble() * 2 - 1)).ToString("G3"),
                        Asymmetry = (2 + 10 * (rnd.NextDouble() * 2 - 1)).ToString("G3"),
                        Frequency = (50 + 0.1 * (rnd.NextDouble() * 2 - 1)).ToString("F2"),
                        CycleResult = resultTable
                    };
                    resultTable.Results.Add(resultRow);

                    resultRow.RatioErrorSamples = "";
                    resultRow.PhaseDispSamples = "";
                    for (int rs = 1; rs <= 10; rs++)
                    {
                        resultRow.RatioErrorSamples += (0.2 + 0.5 * (rnd.NextDouble() * 2 - 1)).ToString("G3") + ";";
                        resultRow.PhaseDispSamples += (2 + 10 * (rnd.NextDouble() * 2 - 1)).ToString("G3") + ";";
                    }
                }
            }

            return report;
        }
    }
}
