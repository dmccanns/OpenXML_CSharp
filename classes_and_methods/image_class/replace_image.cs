using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using PresentationNs = DocumentFormat.OpenXml.Presentation;
using DrawingNs = DocumentFormat.OpenXml.Drawing;

//reference:
//https://social.technet.microsoft.com/wiki/contents/articles/17967.processing-power-point-templates-using-openxml.aspx
public partial class Image
{
    public static void ReplaceImage(string file, string replacementImagePath)
    {
        using (var presentation = PresentationDocument.Open(file, true))
        {
            var slideParts = presentation
                .PresentationPart
                !.SlideParts;

            List<PresentationNs.Picture>? totalElements = new List<PresentationNs.Picture>();
            //iterate through each slide
            int slideIndex = 0;
            foreach (var slidePart in slideParts)
            {
                //gets the tree for that slide
                var tree = slidePart
                .Slide
                .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>()
                .First();

                //gets picture elements on that slide
                List<PresentationNs.Picture>? elements = tree.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Where(
                    p =>
                    {
                        // select non visual picture elements
                        var nvp = p.GetFirstChild<PresentationNs.NonVisualPictureProperties>();
                        if (nvp != null)
                        {
                            var nvd = nvp.GetFirstChild<PresentationNs.NonVisualDrawingProperties>();
                            if (nvd != null)
                            {
                                var extList = nvd.GetFirstChild<DrawingNs.NonVisualDrawingPropertiesExtensionList>();
                                if (extList != null)
                                {
                                    var extensions = extList.Descendants<DrawingNs.NonVisualDrawingPropertiesExtension>().Where(
                                        e => e.Uri == "{generated-asset}"
                                    );
                                    if (extensions.Count() > 0)
                                    {
                                        //TODO: add them to a list or something
                                        foreach (var extension in extensions)
                                        {
                                            var tags = extension.Descendants();
                                            foreach (var tag in tags)
                                            {
                                                // Console.WriteLine("Tag Pair");
                                                // Console.WriteLine(tag.LocalName);
                                                // Console.WriteLine(tag.InnerText);
                                                // Console.WriteLine(' ');
                                            }
                                        }
                                        //add this element to the picture list
                                        return true;
                                    }

                                }
                            }
                        }
                        return false;
                    }
                ).ToList();

                totalElements.AddRange(elements);
                slideIndex++;
            }

            //check if any elements were found
            if (!totalElements.Any())
            {
                Console.WriteLine("No Generated Pictures found");
                return;
            }

            //iterate through elements, find blip and replace the Embed id
            //has to be done by slides since reference is added to slidePart
            foreach (var slidePart in slideParts)
            {
                var tree = slidePart
                .Slide
                .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>()
                .First();

                //get pictures on the slide that are generated Assets
                List<PresentationNs.Picture>? slideElements = tree.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Where(p => totalElements.Contains(p)).ToList();

                foreach (PresentationNs.Picture picture in slideElements)
                {
                    //add image to document
                    var imagePart = slidePart.AddImagePart(ImagePartType.Png);
                    //TODO: change to list of images so each can be different
                    using (var stream = File.OpenRead(replacementImagePath))
                    {
                        imagePart.FeedData(stream);
                    }
                    // get relationship ID
                    var relID = slidePart.GetIdOfPart(imagePart);

                    //get the blip element and change it's embed id, changing the picture
                    var blip = picture.Descendants<DrawingNs.Blip>().First();
                    if (blip != null)
                    {
                        blip.Embed = relID;
                    }
                    else
                    {
                        Console.WriteLine("Picture formatted incorrectly");
                    }
                }

            }
        }
    }

}