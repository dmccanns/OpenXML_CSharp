using PackagingNs = DocumentFormat.OpenXml.Packaging;
using PresentationNs = DocumentFormat.OpenXml.Presentation;
using DrawingNs = DocumentFormat.OpenXml.Drawing;
//^namespaces defined for clarity and ease of reading

public partial class Image
{
    public static void AddImage(string file, string imagePath,  PropertyTag[] properties)
    {
        using (var presentation = PackagingNs.PresentationDocument.Open(file, true))
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
                //TODO: find out which is first so images can be added to the first slide
            }
            var slidePart = slideParts.ElementAt(1);

            //begin image creation code - taken from online 
            //reference: https://stackoverflow.com/questions/35361079/how-i-add-image-in-powerpoint-with-openxml-c-sharp
            var part = slidePart
                .AddImagePart(PackagingNs.ImagePartType.Png);
            using (var stream = File.OpenRead(imagePath))
            {
                part.FeedData(stream);
            }
            var tree = slidePart
                .Slide
                .Descendants<PresentationNs.ShapeTree>()
                .First();
            var picture = new PresentationNs.Picture();
            picture.NonVisualPictureProperties = new PresentationNs.NonVisualPictureProperties();
            var drawingProperties = new PresentationNs.NonVisualDrawingProperties
            {
                Name = "Generated Shape",
                Id = (UInt32)tree.ChildElements.Count - 1,
            };
            picture.NonVisualPictureProperties.Append(drawingProperties);

            var nonVisualPictureDrawingProperties = new PresentationNs.NonVisualPictureDrawingProperties();
            nonVisualPictureDrawingProperties.Append(new DrawingNs.PictureLocks()
            {
                NoChangeAspect = true
            });
            picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
            picture.NonVisualPictureProperties.Append(new PresentationNs.ApplicationNonVisualDrawingProperties());

            var blipFill = new PresentationNs.BlipFill();
            var blip1 = new DrawingNs.Blip()
            {
                Embed = slidePart.GetIdOfPart(part)
            };
            var blipExtensionList1 = new DrawingNs.BlipExtensionList();
            var blipExtension1 = new DrawingNs.BlipExtension()
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
            var stretch = new DrawingNs.Stretch();
            stretch.Append(new DrawingNs.FillRectangle());
            blipFill.Append(blip1);
            blipFill.Append(stretch);
            picture.Append(blipFill);

            picture.ShapeProperties = new PresentationNs.ShapeProperties();
            picture.ShapeProperties.Transform2D = new DrawingNs.Transform2D();
            picture.ShapeProperties.Transform2D.Append(new DrawingNs.Offset
            {
                X = 0,
                Y = 0,
            });
            picture.ShapeProperties.Transform2D.Append(new DrawingNs.Extents
            {
                Cx = 1000000,
                Cy = 1000000,
            });
            picture.ShapeProperties.Append(new DrawingNs.PresetGeometry
            {
                Preset = DrawingNs.ShapeTypeValues.Rectangle
            });

        
            /* 
            Add Blip Extension to drawing Properties (p:cNvPr).
            Blip Extension is used since it persists(idk why, just trial and error)
            */
            var blipExtensionList2 = new DrawingNs.BlipExtensionList();
            var blipExtension2 = new DrawingNs.BlipExtension()
            {
                Uri = "{generated-asset}"
            };
            blipExtension2.InnerXml = string.Join(' ', properties.Select((p)=> $"<{p.tagName} xmlns=\"\">{p.tagValue}</{p.tagName}>"));


            blipExtensionList2.Append(blipExtension2);
            drawingProperties.Append(blipExtensionList2);

            tree.Append(picture);
        }
    }
}