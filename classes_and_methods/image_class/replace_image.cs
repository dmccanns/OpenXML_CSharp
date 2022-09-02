using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
public partial class Image
{
    public static void ReplaceImage(string file, string tagName)
    {
        using (var presentation = PresentationDocument.Open(file, true))
        {
            var slideParts = presentation
                .PresentationPart
                !.SlideParts;

            //iterate through each slide
            foreach (var slidePart in slideParts)
            {
                //gets the tree for that slide
                var tree = slidePart
                .Slide
                .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>()
                .First();

                //gets picture elements on that slide
                List<Picture> elements = tree.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Where(
                    p=>{
                        try {
                            p.GetAttribute("Type", "");
                        }catch(KeyNotFoundException ex){
                            Console.WriteLine("ex:" + ex);
                            return false;
                        }
                        //if it doesn't fail, then just return true
                        return true;
                        //select non visual picture elements
                        // var nvp = p.GetFirstChild<DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties>();
                        // if (nvp !=null){
                        //     //if they have a Tag List
                        //     var descendants = nvp.Descendants();
                        //     foreach (var descendant in descendants){
                        //         Console.WriteLine(descendant);
                        //         Console.WriteLine(' ');
                        //     }
                        //     return nvp.GetFirstChild<TagList>() != null ? true: false;
                        // }
                        // return false;
                    }
                ).ToList();

                // ?.GetFirstChild<Tag>?.Val?.Value?.Contains("SHAPETAG") ?? false

                foreach (var element in elements)
                {
                    Console.WriteLine(element.InnerXml);
                }
                //filter child elements for pictures that match
                // List<Picture> generatedPictures = tree.ChildElements.Where
                // (p => p is Picture);

                //r.GetFirstChild<Tag>().Val.Value.Contains("tagname")
            }

        }
    }

}