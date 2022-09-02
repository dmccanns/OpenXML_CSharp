using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
class CustomTag{
    public static void AddCustomTagToPicture(string propertyName, string propertyValue, PresentationPart presentationPart , Picture picture){
        
    //creates the tag.xml file
    TagList tagList = new TagList();
    Tag tag = new Tag() { Name = propertyName, Val = propertyValue };
    tagList.Append(tag);

    //Creates new part w/ new id
    UserDefinedTagsPart userTagsPart = presentationPart.AddNewPart<UserDefinedTagsPart>();
    userTagsPart.TagList = tagList;

    //get the nvPicPr child of the picture element
    NonVisualPictureProperties? pictureProperties = picture.GetFirstChild<NonVisualPictureProperties>();
    if (pictureProperties == null){
        Console.WriteLine("XML formatting wrong");
        return;
    }

    var custData = presentationPart.AddCustomXmlPart("CustomerData");
    //get the nvPr child if it exists - if not make a new one and append it
    ApplicationNonVisualDrawingProperties? drawingProperties = pictureProperties.GetFirstChild<ApplicationNonVisualDrawingProperties>();
    if (drawingProperties == null){
        ApplicationNonVisualDrawingProperties newDrawingProperties = new ApplicationNonVisualDrawingProperties();
        pictureProperties.Append(newDrawingProperties);
        drawingProperties = pictureProperties.GetFirstChild<ApplicationNonVisualDrawingProperties>();
    }
    
    //customer tags doesn't exist since this is a new image
    var dataListElement = new CustomerDataList();

    //create the reference to the tags and append it in the pictureProperties element
    var id = presentationPart.GetIdOfPart(userTagsPart);
    var refTag = new CustomerDataTags(){
        Id = id
    };
    dataListElement.Append(refTag);
    drawingProperties!.Append(dataListElement);
    //can't be null due to if statement on line 25
    }
}