using PackagingObj = DocumentFormat.OpenXml.Packaging;
using PresentationObj = DocumentFormat.OpenXml.Presentation;
using DrawingObj = DocumentFormat.OpenXml.Drawing;
public partial class Image
{
    public static void AddImage(string file, string image)
    {
        using (var presentation = PackagingObj.PresentationDocument.Open(file, true))
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
                .AddImagePart(PackagingObj.ImagePartType.Png);
            using (var stream = File.OpenRead(image))
            {
                part.FeedData(stream);
            }
            var tree = slidePart
                .Slide
                .Descendants<PresentationObj.ShapeTree>()
                .First();
            var picture = new PresentationObj.Picture();
            picture.NonVisualPictureProperties = new PresentationObj.NonVisualPictureProperties();
            var drawingProperties = new PresentationObj.NonVisualDrawingProperties
            {
                Name = "Generated Shape",
                Id = (UInt32)tree.ChildElements.Count - 1,
            };
            picture.NonVisualPictureProperties.Append(drawingProperties);

            var nonVisualPictureDrawingProperties = new PresentationObj.NonVisualPictureDrawingProperties();
            nonVisualPictureDrawingProperties.Append(new DrawingObj.PictureLocks()
            {
                NoChangeAspect = true
            });
            picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
            picture.NonVisualPictureProperties.Append(new PresentationObj.ApplicationNonVisualDrawingProperties());

            var blipFill = new PresentationObj.BlipFill();
            var blip1 = new DrawingObj.Blip()
            {
                Embed = slidePart.GetIdOfPart(part)
            };
            var blipExtensionList1 = new DrawingObj.BlipExtensionList();
            var blipExtension1 = new DrawingObj.BlipExtension()
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
            var stretch = new DrawingObj.Stretch();
            stretch.Append(new DrawingObj.FillRectangle());
            blipFill.Append(blip1);
            blipFill.Append(stretch);
            picture.Append(blipFill);

            picture.ShapeProperties = new PresentationObj.ShapeProperties();
            picture.ShapeProperties.Transform2D = new DrawingObj.Transform2D();
            picture.ShapeProperties.Transform2D.Append(new DrawingObj.Offset
            {
                X = 0,
                Y = 0,
            });
            picture.ShapeProperties.Transform2D.Append(new DrawingObj.Extents
            {
                Cx = 1000000,
                Cy = 1000000,
            });
            picture.ShapeProperties.Append(new DrawingObj.PresetGeometry
            {
                Preset = DrawingObj.ShapeTypeValues.Rectangle
            });

        
            /* 
            Add Blip Extension to drawing Properties (p:cNvPr).
            Blip Extension is used since it persists(idk why, just trial and error)
            */
            var blipExtensionList2 = new DrawingObj.BlipExtensionList();
            var blipExtension2 = new DrawingObj.BlipExtension()
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