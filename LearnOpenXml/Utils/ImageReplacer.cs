using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace LearnOpenXml.Utils;

// https://learn.microsoft.com/en-us/office/open-xml/how-to-insert-a-picture-into-a-word-processing-document
public class ImageReplacer
{
    public static void InsertPicture(WordprocessingDocument wordProcessingDocument, OpenXmlElement imagePlaceholder, Stream image)
    {
        var mainPart = wordProcessingDocument.MainDocumentPart!;
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        
        imagePart.FeedData(image);

        AddImageToBody(imagePlaceholder, mainPart.GetIdOfPart(imagePart));
    }
    
    private static void AddImageToBody(OpenXmlElement imagePlaceholder, string relationshipId)
    {
        // Define the reference of the image.
        var element =
            new Drawing(new DW.Inline(new DW.Extent
                    { Cx = 1280000, Cy = 1280000 },
                new DW.EffectExtent()
                {
                    LeftEdge = 0L, TopEdge = 0L,
                    RightEdge = 0L, BottomEdge = 0L
                },
                new DW.DocProperties
                {
                    Id = (UInt32Value)1U,
                    Name = "logo"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                new A.Graphic(new A.GraphicData(new PIC.Picture(new PIC.NonVisualPictureProperties(new PIC.NonVisualDrawingProperties()
                        {
                            Id = (UInt32Value)0U,
                            Name = "logo.png"
                        },
                        new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(new A.Blip(new A.BlipExtensionList(new A.BlipExtension()
                        {
                            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                        }))
                        {
                            Embed = relationshipId,
                            CompressionState =
                                A.BlipCompressionValues.Print
                        },
                        new A.Stretch(new A.FillRectangle())),
                    new PIC.ShapeProperties(new A.Transform2D(new A.Offset() { X = 0L, Y = 0L },
                            new A.Extents() { Cx = 1280000, Cy = 1280000 }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                EditId = "50D07946"
            });

        imagePlaceholder.Parent!.ReplaceChild(new Paragraph(new Run(element)), imagePlaceholder);
    }
}