using System.IO;
using System.Linq;
using System.Windows;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using OpenXmlPowerTools;

namespace WordReportTest.Export
{
    public class WordExportImage
    {
        //private static object _imageObject;
        //public static void SetImageObject(object obj)
        //{
        //    _imageObject = obj;
        //}

        //public static void InsertPicture(string document)
        //{
        //    var drawing = Utility.GetDrawingFromXaml(_imageObject);

        //    var bounds = drawing.Bounds;

        //    var wmfStream = new MemoryStream();

        //    using (var g = CreateEmf(wmfStream, bounds))
        //        Utility.RenderDrawingToGraphics(drawing, g);

        //    wmfStream.Position = 0;


        //    using (WordprocessingDocument wordprocessingDocument =
        //        WordprocessingDocument.Open(document, true))
        //    {
        //        MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

        //        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Emf);

        //        imagePart.FeedData(wmfStream);

        //        AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
        //    }
        //}


        private static XElement GetImageFromXaml(WordprocessingDocument wordDoc, object xamlImage)
        {
            var drawing = Utility.GetDrawingFromXaml(xamlImage);
            var bounds = drawing.Bounds;
            var wmfStream = new MemoryStream();

            using (var g = CreateEmf(wmfStream, bounds))
                Utility.RenderDrawingToGraphics(drawing, g);
            wmfStream.Position = 0;

            var mainPart = wordDoc.MainDocumentPart;
            var imagePart = mainPart.AddImagePart(ImagePartType.Emf);
            imagePart.FeedData(wmfStream);

            string relationshipId = mainPart.GetIdOfPart(imagePart);

            // Define the reference of the image.
            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() {Cx = 6382385L, Cy = 2400935L},
                        new DW.EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties()
                        {
                            Id = (UInt32Value) 1U,
                            Name = "Picture 1"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() {NoChangeAspect = true}),
                        new A.Graphic(
                            new A.GraphicData(
                                    new PIC.Picture(
                                        new PIC.NonVisualPictureProperties(
                                            new PIC.NonVisualDrawingProperties()
                                            {
                                                Id = (UInt32Value) 0U,
                                                Name = "NewImage.Wmf"
                                            },
                                            new PIC.NonVisualPictureDrawingProperties()),
                                        new PIC.BlipFill(
                                            new A.Blip(
                                                new A.BlipExtensionList(
                                                    new A.BlipExtension()
                                                    {
                                                        Uri =
                                                            "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                    })
                                            )
                                            {
                                                Embed = relationshipId,
                                                CompressionState =
                                                    A.BlipCompressionValues.Print
                                            },
                                            new A.Stretch(
                                                new A.FillRectangle())),
                                        new PIC.ShapeProperties(
                                            new A.Transform2D(
                                                new A.Offset() {X = 0L, Y = 0L},
                                                new A.Extents() {Cx = 6382385L, Cy = 2400935L}),
                                            new A.PresetGeometry(
                                                    new A.AdjustValueList()
                                                )
                                                {Preset = A.ShapeTypeValues.Rectangle}))
                                )
                                {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
                    )
                    {
                        DistanceFromTop = (UInt32Value) 0U,
                        DistanceFromBottom = (UInt32Value) 0U,
                        DistanceFromLeft = (UInt32Value) 0U,
                        DistanceFromRight = (UInt32Value) 0U,
                        EditId = "50D07946"
                    });

            return new XElement((new Run(element)).OuterXml);
        }

        public static System.Drawing.Graphics CreateEmf(Stream wmfStream, Rect bounds)
        {
            if (bounds.Width == 0 || bounds.Height == 0) bounds = new Rect(0, 0, 1, 1);
            using (System.Drawing.Graphics refDC = System.Drawing.Graphics.FromImage(new System.Drawing.Bitmap(1, 1)))
            {
                System.Drawing.Graphics graphics =
                    System.Drawing.Graphics.FromImage(
                        new System.Drawing.Imaging.Metafile(wmfStream, refDC.GetHdc(),
                            bounds.ToGdiPlus(),
                            System.Drawing.Imaging.MetafileFrameUnit.Pixel,
                            System.Drawing.Imaging.EmfType.EmfPlusDual));

                return graphics;
            }
        }


        //var paragraphs = wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>().ToList();

        //foreach (var p in paragraphs)
        //{
        //    var contents = p.Descendants<Text>().Select(t => t.Text).StringConcatenate();
        //    if (contents.Contains("LineChartTest"))
        //    {
        //        foreach (var run in p.Descendants<Run>().ToList())
        //            run.Remove();
        //        p.AppendChild(new Run(element));
        //    }
        //}

        // Append the reference to body, the element should be in a Run.
        //wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));

        //}


        //public static void InsertAPicture(string document, string fileName)
        //{
        //    using (WordprocessingDocument wordprocessingDocument =
        //        WordprocessingDocument.Open(document, true))
        //    {
        //        MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

        //        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

        //        using (FileStream stream = new FileStream(fileName, FileMode.Open))
        //        {
        //            imagePart.FeedData(stream);
        //        }

        //        AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
        //    }
        //}

        //private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        //{
        //    // Define the reference of the image.
        //    var element =
        //         new Drawing(
        //             new DW.Inline(
        //                 new DW.Extent() { Cx = 6382385L, Cy = 2400935L },
        //                 new DW.EffectExtent()
        //                 {
        //                     LeftEdge = 0L,
        //                     TopEdge = 0L,
        //                     RightEdge = 0L,
        //                     BottomEdge = 0L
        //                 },
        //                 new DW.DocProperties()
        //                 {
        //                     Id = (UInt32Value)1U,
        //                     Name = "Picture 1"
        //                 },
        //                 new DW.NonVisualGraphicFrameDrawingProperties(
        //                     new A.GraphicFrameLocks() { NoChangeAspect = true }),
        //                 new A.Graphic(
        //                     new A.GraphicData(
        //                         new PIC.Picture(
        //                             new PIC.NonVisualPictureProperties(
        //                                 new PIC.NonVisualDrawingProperties()
        //                                 {
        //                                     Id = (UInt32Value)0U,
        //                                     Name = "New Bitmap Image.Wmf"
        //                                 },
        //                                 new PIC.NonVisualPictureDrawingProperties()),
        //                             new PIC.BlipFill(
        //                                 new A.Blip(
        //                                     new A.BlipExtensionList(
        //                                         new A.BlipExtension()
        //                                         {
        //                                             Uri =
        //                                                "{28A0092B-C50C-407E-A947-70E740481C1C}"
        //                                         })
        //                                 )
        //                                 {
        //                                     Embed = relationshipId,
        //                                     CompressionState =
        //                                     A.BlipCompressionValues.Print
        //                                 },
        //                                 new A.Stretch(
        //                                     new A.FillRectangle())),
        //                             new PIC.ShapeProperties(
        //                                 new A.Transform2D(
        //                                     new A.Offset() { X = 0L, Y = 0L },
        //                                     new A.Extents() { Cx = 6382385L, Cy = 2400935L }),
        //                                 new A.PresetGeometry(
        //                                     new A.AdjustValueList()
        //                                 )
        //                                 { Preset = A.ShapeTypeValues.Rectangle }))
        //                     )
        //                     { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
        //             )
        //             {
        //                 DistanceFromTop = (UInt32Value)0U,
        //                 DistanceFromBottom = (UInt32Value)0U,
        //                 DistanceFromLeft = (UInt32Value)0U,
        //                 DistanceFromRight = (UInt32Value)0U,
        //                 EditId = "50D07946"
        //             });


        //    //var paragraphs = wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>().ToList();

        //    //foreach (var p in paragraphs)
        //    //{
        //    //    var contents = p.Descendants<Text>().Select(t => t.Text).StringConcatenate();
        //    //    if (contents.Contains("LineChartTest"))
        //    //    {
        //    //        foreach (var run in p.Descendants<Run>().ToList())
        //    //            run.Remove();
        //    //        p.AppendChild(new Run(element));
        //    //    }
        //    //}

        //    // Append the reference to body, the element should be in a Run.
        //    //wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));

        //}
    }
}
