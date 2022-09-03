using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
public partial class Image
{
    public static void AddImage(string file, string image)
    {
        using (var presentation = PresentationDocument.Open(file, true))
        {   
            var presentationPart = presentation.PresentationPart; 
            if (presentationPart == null){
                Console.WriteLine("null presentation");
                return;
            }
            var slideParts = presentationPart
                .SlideParts;

            foreach (var slidePart1 in slideParts)
            {
                Console.WriteLine(slidePart1);
            }
            var slidePart = slideParts.ElementAt(1);

            var part = slidePart
                .AddImagePart(ImagePartType.Png);

            using (var stream = File.OpenRead(image))
            {
                part.FeedData(stream);
            }

            var tree = slidePart
                .Slide
                .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>()
                .First();
            
            var picture = new DocumentFormat.OpenXml.Presentation.Picture();

            picture.NonVisualPictureProperties = new NonVisualPictureProperties();
            var drawingProperties = new NonVisualDrawingProperties
            {
                Name = "My Shape",
                Id = (UInt32)tree.ChildElements.Count - 1,
            };

            picture.NonVisualPictureProperties.Append(drawingProperties);

            var nonVisualPictureDrawingProperties = new NonVisualPictureDrawingProperties();
            nonVisualPictureDrawingProperties.Append(new DocumentFormat.OpenXml.Drawing.PictureLocks()
            {
                NoChangeAspect = true
            });
            picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
            picture.NonVisualPictureProperties.Append(new ApplicationNonVisualDrawingProperties());

            var blipFill = new DocumentFormat.OpenXml.Presentation.BlipFill();
            var blip1 = new DocumentFormat.OpenXml.Drawing.Blip()
            {
                Embed = slidePart.GetIdOfPart(part)
            };
            var blipExtensionList1 = new DocumentFormat.OpenXml.Drawing.BlipExtensionList();
            var blipExtension1 = new DocumentFormat.OpenXml.Drawing.BlipExtension()
            {
                Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
            };
            var useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi()
            {
                Val = false
            };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            blipExtension1.Append(useLocalDpi1);
            blipExtensionList1.Append(blipExtension1);
            blip1.Append(blipExtensionList1);
            var stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
            stretch.Append(new DocumentFormat.OpenXml.Drawing.FillRectangle());
            blipFill.Append(blip1);
            blipFill.Append(stretch);
            picture.Append(blipFill);

            picture.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();
            picture.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();
            picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Offset
            {
                X = 0,
                Y = 0,
            });
            picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
            {
                Cx = 1000000,
                Cy = 1000000,
            });
            picture.ShapeProperties.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry
            {
                Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
            });

        
            /* 
            Add Blip Extension to drawing Properties (p:cNvPr).
            Blip Extension is used since it persists(idk why, just trial and error)
            */
            var blipExtensionList2 = new DocumentFormat.OpenXml.Drawing.BlipExtensionList();
            var blipExtension2 = new DocumentFormat.OpenXml.Drawing.BlipExtension()
            {
                Uri = "{generated-asset}"
            };
            blipExtension2.InnerXml = "<Type xmlns=\"\">line-graph</Type>";
            blipExtensionList2.Append(blipExtension2);
            drawingProperties.Append(blipExtensionList2);

            tree.Append(picture);
        }
    }
}