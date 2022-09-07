using PackagingNs = DocumentFormat.OpenXml.Packaging;
using PresentationNs = DocumentFormat.OpenXml.Presentation;
using DrawingNs = DocumentFormat.OpenXml.Drawing;
//general methods
public class General
{
    //returns an ordered list of slides for a given presentation
    List<PackagingNs.SlidePart> getOrderedSlides(PackagingNs.PresentationPart presentationPart)
    {
        List<PackagingNs.SlidePart> orderedSlideList = new List<PackagingNs.SlidePart>();

        //Get slideIdList - it should never be null
        var slideIds = presentationPart.Presentation?.SlideIdList;
        if (slideIds == null)
        {
            Console.WriteLine("Your file is messed up");
            throw new Exception();
        }

        foreach (PresentationNs.SlideId slideId in slideIds)
        {
            // get the Part by its RelationshipId - relationshipId is mandatory so can't be null
            var slidePart = presentationPart.GetPartById(slideId.RelationshipId!);

            //again, if slidePart is not a SlidePart the ppt file is messed up
            if (slidePart is PackagingNs.SlidePart)
            {
                //casting it since it was giving me an annoying warning. Could def be rewritten
                var slidePartCast = slidePart as PackagingNs.SlidePart;
                orderedSlideList.Add(slidePartCast!);
            }
            else
            {
                Console.WriteLine("Your file is messed up");
                throw new Exception();
            }
        }
        return orderedSlideList;
    }
}